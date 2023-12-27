using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace TyBus_Intranet_Test_V3
{
    public partial class TourCarIncome : System.Web.UI.Page
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
        DataTable tbResult = new DataTable();

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
                string vSQLStr = "";
                string vTempResult = "";

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }

                    if (!IsPostBack)
                    {
                        //每車年強制險保費 CarYearInsurance
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CarYearInsu'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        CarYearInsurance.Text = (vTempResult != "") ? vTempResult : "0";
                        //每車年第三責任險 CarThirdInsurance
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CarThirdInsu'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        CarThirdInsurance.Text = (vTempResult != "") ? vTempResult : "0";
                        //駕駛員每人月勞保費 DriverLiMonth
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_DriverLi'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        DriverLiMonth.Text = (vTempResult != "") ? vTempResult : "0";
                        //駕駛員每人月健保費 DriverHiMonth
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_DriverHi'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        DriverHiMonth.Text = (vTempResult != "") ? vTempResult : "0";
                        //駕駛員獎金每人每月 DriverBoundsMonth
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_DriverBounds'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        DriverBoundsMonth.Text = (vTempResult != "") ? vTempResult : "0";
                        //賠償費 CompensateLoss
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CompLoss11'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eCompensateLoss11.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CompLoss13'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eCompensateLoss13.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CompLoss15'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eCompensateLoss15.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CompLoss16'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eCompensateLoss16.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CompLoss17'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eCompensateLoss17.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CompLoss18'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eCompensateLoss18.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CompLoss19'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eCompensateLoss19.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CompLoss20'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eCompensateLoss20.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CompLoss22'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eCompensateLoss22.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CompLoss23'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eCompensateLoss23.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CompLoss25'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eCompensateLoss25.Text = (vTempResult != "") ? vTempResult : "0";
                        //專車收入 TranIncome
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_TranIncome11'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eTranIncome11.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_TranIncome13'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eTranIncome13.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_TranIncome15'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eTranIncome15.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_TranIncome16'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eTranIncome16.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_TranIncome17'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eTranIncome17.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_TranIncome18'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eTranIncome18.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_TranIncome19'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eTranIncome19.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_TranIncome20'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eTranIncome20.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_TranIncome22'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eTranIncome22.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_TranIncome23'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eTranIncome23.Text = (vTempResult != "") ? vTempResult : "0";
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_TranIncome25'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        eTranIncome25.Text = (vTempResult != "") ? vTempResult : "0";
                        //服裝費 ClothCost
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_ClothCost'";
                        vTempResult = PF.GetValue(vConnStr, vSQLStr, "Content");
                        ClothCost.Text = (vTempResult != "") ? vTempResult : "0";
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

        private void SaveParaValues(string fItemName, string fContentValues)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select count(Content) RCount from SysFlag where FormName = 'unSysFlag' and ControlItem = '" + fItemName + "' ";
            string vRCount = PF.GetValue(vConnStr, vSQLStr, "RCount");
            vSQLStr = (vRCount == "0") ?
                      "insert into SysFlag (FormName, ControlItem, Content) values('unSysFlag', '" + fItemName + "', '" + fContentValues + "')" :
                      "update SysFlag set Content = '" + fContentValues + "' where FormName = 'unSysFlag' and ControlItem = '" + fItemName + "' ";
            PF.ExecSQL(vConnStr, vSQLStr);
        }

        /// <summary>
        /// 開始計算
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSearch_Click(object sender, EventArgs e)
        {
            string vYear_C = "";
            string vMonth = "";
            string vYM_C = "";

            string vBuDate_S = "";
            string vBuDate_E = "";
            string vRowsName = "";
            string vSelectStr_Temp = "";
            //string vTempStr = "";
            //double vTempFloat = 0.0;
            //double vTempTotalOut = 0.0;
            //double vTempTotalInOut = 0.0;
            //======================================================================================================

            string vCalDepNo = "";
            string vInOutFieldName = "";
            string vCashFieldName = "";
            //======================================================================================================
            double vTourIncome = 0.0; //遊覽車收入
            double vTranIncome = 0.0; //交通車收入
            double vBusIncome = 0.0; //班公車收入
            double vTotalIncome = 0.0; //營業收入合計
            double vDieselAMT = 0.0; //柴油費
            double vEngineOilAMT = 0.0; //機油費
            double vOtherOil = 0.0; //其他油料費
            double vFixAMT = 0.0; //修車材料費
            double vTireAMT = 0.0; //輪胎費
            double vDepAMT = 0.0; //車輛折舊
            double vCarInst = 0.0; //車輛保險費
            double vCompensateLoss = 0.0; //車輛賠償費
            double vLicenceAMT = 0.0; //證照費
            double vDriverPayment = 0.0; //駕駛員
            double vBoundsFee = 0.0; //獎金提列
            double vRetirePay = 0.0; //退休金
            double vLiFee = 0.0; //勞保費
            double vHiFee = 0.0; //健保費
            double vFixEmp = 0.0; //修車員工
            double vClothCost = 0.0; //服裝費用
            double vTotalOutpay = 0.0; //支出合計
            double vTotalInOut = 0.0; //本期損益
            double vTourEmp = 0.0; //駕駛員人數基準數
            double vTourCount = 0.0; //使用車輛數基準數
            double vCarCount = 0.0; //站實際車輛數
            double vTourKM = 0.0; //遊覽車里程
            double vTransKM = 0.0; //交通車里程
            double vBus_KM = 0.0; //班公車里程
            double vNoneKM = 0.0; //非營運里程
            double vTotalKMs = 0.0; //里程合計
            //double vCarIncome = 0.0; //每車收入
            //double vCarOutpay = 0.0; //每車支出
            //double vCarInOut = 0.0; //每車盈餘
            //======================================================================================================

            List<string> ItemList = new List<string>();
            ItemList.Add("TourIncome");
            ItemList.Add("TranIncome");
            ItemList.Add("BusIncome");
            ItemList.Add("TotalIncome");
            ItemList.Add("DieselAMT");
            ItemList.Add("EngineOilAMT");
            ItemList.Add("OtherOil");
            ItemList.Add("FixAMT");
            ItemList.Add("TireAMT");
            ItemList.Add("DepAMT");
            ItemList.Add("CarInst");
            ItemList.Add("CompensateLoss");
            ItemList.Add("LicenceAMT");
            ItemList.Add("DriverPayment");
            ItemList.Add("BoundsFee");
            ItemList.Add("RetirePay");
            ItemList.Add("LiFee");
            ItemList.Add("HiFee");
            ItemList.Add("FixEmp");
            ItemList.Add("ClothCost");
            ItemList.Add("TotalOutpay");
            ItemList.Add("TotalInOut");

            string vSQLStr = "select DepNo from Department where isnull(IsStation, 'X') = 'V' and isnull(InSHReport, 'X') = 'V' order by DepNo";
            SqlConnection connSearchDepNo = new SqlConnection(PF.GetConnectionStr(Request.ApplicationPath));
            SqlCommand cmdSearchDepNo = new SqlCommand(vSQLStr, connSearchDepNo);
            connSearchDepNo.Open();
            SqlDataReader drSearchDepNo = cmdSearchDepNo.ExecuteReader();
            List<string> DepList = new List<string>();
            while (drSearchDepNo.Read())
            {
                DepList.Add(drSearchDepNo["DepNo"].ToString().Trim());
            }
            drSearchDepNo.Close();
            cmdSearchDepNo.Cancel();
            connSearchDepNo.Close();

            if (DriveYM.Text.Trim().Length == 6)
            {
                //先把畫面上的幾個參數值存回資料庫
                //每車年強制險保費 CarYearInsurance
                SaveParaValues("a_CarYearInsu", CarYearInsurance.Text.Trim());
                //每車年第三責任險 CarThirdInsurance
                SaveParaValues("a_CarThirdInsu", CarThirdInsurance.Text.Trim());
                //駕駛員每人月勞保費 DriverLiMonth
                SaveParaValues("a_DriverLi", DriverLiMonth.Text.Trim());
                //駕駛員每人月健保費 DriverHiMonth
                SaveParaValues("a_DriverHi", DriverHiMonth.Text.Trim());
                //駕駛員獎金每人每月 DriverBoundsMonth
                SaveParaValues("a_DriverBounds", DriverBoundsMonth.Text.Trim());
                //賠償費 CompensateLoss
                SaveParaValues("a_CompLoss11", eCompensateLoss11.Text.Trim());
                SaveParaValues("a_CompLoss13", eCompensateLoss13.Text.Trim());
                SaveParaValues("a_CompLoss15", eCompensateLoss15.Text.Trim());
                SaveParaValues("a_CompLoss16", eCompensateLoss16.Text.Trim());
                SaveParaValues("a_CompLoss17", eCompensateLoss17.Text.Trim());
                SaveParaValues("a_CompLoss18", eCompensateLoss18.Text.Trim());
                SaveParaValues("a_CompLoss19", eCompensateLoss19.Text.Trim());
                SaveParaValues("a_CompLoss20", eCompensateLoss20.Text.Trim());
                SaveParaValues("a_CompLoss22", eCompensateLoss22.Text.Trim());
                SaveParaValues("a_CompLoss23", eCompensateLoss23.Text.Trim());
                SaveParaValues("a_CompLoss25", eCompensateLoss25.Text.Trim());
                //專車收入
                SaveParaValues("a_TranIncome11", eTranIncome11.Text.Trim());
                SaveParaValues("a_TranIncome13", eTranIncome13.Text.Trim());
                SaveParaValues("a_TranIncome15", eTranIncome15.Text.Trim());
                SaveParaValues("a_TranIncome16", eTranIncome16.Text.Trim());
                SaveParaValues("a_TranIncome17", eTranIncome17.Text.Trim());
                SaveParaValues("a_TranIncome18", eTranIncome18.Text.Trim());
                SaveParaValues("a_TranIncome19", eTranIncome19.Text.Trim());
                SaveParaValues("a_TranIncome20", eTranIncome20.Text.Trim());
                SaveParaValues("a_TranIncome22", eTranIncome22.Text.Trim());
                SaveParaValues("a_TranIncome23", eTranIncome23.Text.Trim());
                SaveParaValues("a_TranIncome25", eTranIncome25.Text.Trim());
                //服裝費 ClothCost
                SaveParaValues("a_ClothCost", ClothCost.Text.Trim());
                //開始進行計算
                vBuDate_S = DriveYM.Text.Trim();
                vMonth = vBuDate_S.Substring(4, 2);
                vBuDate_S = vBuDate_S.Substring(0, 4) + "/" + vBuDate_S.Substring(4, 2) + "/01";
                vBuDate_E = PF.GetMonthLastDay(DateTime.Parse(vBuDate_S), "B");
                vYear_C = (int.Parse(vBuDate_S.Substring(0, 4)) - 1911).ToString();
                vYM_C = vYear_C + vMonth;
                //組合查詢語法
                vSQLStr = "select x.DepNo, " +
                          "       x.TourIncome, " + //遊覽車收入
                          "       x.TranIncome, " + //交通車收入
                          "       x.BusIncome, " + Environment.NewLine + //班公車收入
                          "       cast(0 as float) TotalIncome, " + Environment.NewLine + //營業收入合計
                          "       x.DieselAMT, " + //柴油費
                          "       x.EngineOilAMT, " + //機油費
                          "       x.OtherOil, " + //其他油料費
                          "       x.FixAMT, " + //修車材料費
                          "       x.TireAMT, " + //輪胎費
                          "       x.DepAMT, " + //車輛折舊
                          "       x.CarInst, " + //車輛保險費
                          "       x.CompensateLoss, " + //車輛賠償費
                          "       x.LicenceAMT, " + //證照費
                          "       x.DriverPayment, " + //駕駛員
                          "       x.BoundsFee, " + //獎金提列
                          "       x.RetirePay, " + //退休金
                          "       x.LiFee, " + //勞保費
                          "       x.HiFee, " + //健保費
                          "       x.FixEmp, " + //修車員工
                          "       x.ClothCost, " + Environment.NewLine + //服裝費用
                          "       cast(0 as float) TotalOutpay, " + Environment.NewLine + //支出合計
                          "       cast(0 as float) TotalInOut, " + Environment.NewLine + //本期損益
                          "       x.TourEmp, " + //駕駛員人數基準數
                          "       x.TourCount, " + //使用車輛數基準數
                          "       x.CarCount, " + //站實際車輛數
                          "       x.TourKM, " + //遊覽車里程
                          "       x.TransKM, " + //交通車里程
                          "       x.Bus_KM, " + //班公車里程
                          "       x.NoneKM, " + Environment.NewLine + //非營運里程
                          "       cast(0 as float) TotalKMs, " + Environment.NewLine + //里程合計
                          "       cast(0 as float) CarIncome, " + Environment.NewLine + //每車收入
                          "       cast(0 as float) CarOutpay, " + Environment.NewLine + //每車支出
                          "       cast(0 as float) CarInOut, " + Environment.NewLine + //每車盈餘
                          "       x.CarYearInsu, " + //
                          "       x.CarThirdInsu, " + //
                          "       x.DriverLi, " + //
                          "       x.DriverHi, " + //
                          "       x.DriverBounds " + Environment.NewLine + //
                          "  from ( " + Environment.NewLine +
                          "        select z.DepNo, sum(z.TourIncome) TourIncome, sum(z.TranIncome_Untax) TranIncome, sum(z.BusIncome) BusIncome, " + Environment.NewLine +
                          "               sum(z.DieselAMT) DieselAMT, sum(z.EngineOilAMT) EngineOilAMT, " + Environment.NewLine +
                          "               cast(isnull((select sum(isnull(Debit, 0) - isnull(Credit, 0)) " + Environment.NewLine +
                          "                              from Account" + Environment.NewLine +
                          "                             where [Subject] like '5130%' and right([Subject], 2) = z.DepNo and [Type] <> '6' " + Environment.NewLine +
                          "                               and Rec_Date between '" + vBuDate_S + "' and '" + vBuDate_E + "'), 0) / z.CarCount * z.TourCount as decimal(10, 0)) OtherOil, " + Environment.NewLine +
                          "               sum(z.FixAMT) FixAMT, sum(z.TireAMT) TireAMT, sum(z.DepAMT) DepAMT, " + Environment.NewLine +
                          "               cast(((" + CarYearInsurance.Text.Trim() + " + " + CarThirdInsurance.Text.Trim() + ") / 12 * z.TourCount) as decimal(10, 0)) CarInst," + Environment.NewLine +
                          "               cast(0 as float) CompensateLoss," + Environment.NewLine +
                          "               sum(z.LicenceAMT) LicenceAMT," + Environment.NewLine +
                          "               sum(z.RealCash) DriverPayment, " + Environment.NewLine +
                          "               cast((" + DriverBoundsMonth.Text.Trim() + " * z.TourCount) as decimal(10, 0)) BoundsFee, " + Environment.NewLine +
                          "               sum(z.RetirePay) RetirePay, " + Environment.NewLine +
                          "               cast((" + DriverLiMonth.Text.Trim() + " * z.TourCount) as decimal(10, 0)) LiFee, " + Environment.NewLine +
                          "               cast((" + DriverHiMonth.Text.Trim() + " * z.TourCount) as decimal(10, 0)) HiFee, " + Environment.NewLine +
                          "               cast(0 as float) FixEmp, " + Environment.NewLine +
                          "               cast((" + ClothCost.Text.Trim() + " * z.TourCount) as decimal(10, 0)) ClothCost, " + Environment.NewLine +
                          "               z.TourCount as TourEmp, z.TourCount, z.CarCount, " + Environment.NewLine +
                          "               cast(sum(z.TourKM) as decimal(10,0)) TourKM, cast(sum(z.TransKM) as decimal(10,0)) TransKM, " + Environment.NewLine +
                          "               cast(sum(z.Bus_KM) as decimal(10,0)) Bus_KM, cast(sum(z.NoneKM) as decimal(10,0)) NoneKM, " + Environment.NewLine +
                          "               cast(" + CarYearInsurance.Text.Trim() + " as decimal(10, 0)) CarYearInsu, " + Environment.NewLine +
                          "               cast(" + CarThirdInsurance.Text.Trim() + " as decimal(10, 0)) CarThirdInsu, " + Environment.NewLine +
                          "               cast(" + DriverLiMonth.Text.Trim() + " as decimal(10, 0)) DriverLi, " + Environment.NewLine +
                          "               cast(" + DriverHiMonth.Text.Trim() + " as decimal(10, 0)) DriverHi, " + Environment.NewLine +
                          "               cast(" + DriverBoundsMonth.Text.Trim() + " as decimal(10, 0)) DriverBounds " + Environment.NewLine +
                          "          from ( " + Environment.NewLine +
                          "                select a.DepNo, a.Car_NO, a.Car_ID, " + Environment.NewLine +
                          "                       isnull((select isnull(RealCash, 0) RealCash from PayRec " + Environment.NewLine +
                          "                                where PayDate = '" + vBuDate_S + "'  and PayDur = '1' " + Environment.NewLine +
                          "                                  and EmpNo = (select Driver from Car_infoA where Car_infoA.Car_ID = a.Car_ID and CompanyNo = a.DepNo)), 0) RealCash, " + Environment.NewLine +
                          "                       isnull((select isnull(CashNum74, 0) RetirePay from PayRec " + Environment.NewLine +
                          "                                where PayDate = '" + vBuDate_S + "'  and PayDur = '1' " + Environment.NewLine +
                          "                                  and EmpNo = (select Driver from Car_infoA where Car_infoA.Car_ID = a.Car_ID and CompanyNo = a.DepNo)), 0) RetirePay, " + Environment.NewLine +
                          "                       cast(sum(isnull(a.ACashAMT, 0) + isnull(a.ACashAMT1, 0) + isnull(a.AOAMT, 0) + isnull(a.AOAMT1, 0) + isnull(a.ASPiece, 0) + isnull(a.ASPiece1, 0) + " + Environment.NewLine +
                          "                            isnull(a.ALines, 0) + isnull(a.ALines1, 0)) as decimal(10, 0)) as BusIncome, " + Environment.NewLine +
                          "                       sum(isnull(a.AunTraincome, 0)) as TranIncome_Untax, sum(isnull(a.ARentIncomeB, 0)) as TourIncome, sum(isnull(b.TourCount_OT, 0) + isnull(b.TourCount_ET, 0)) as TourCount, " + Environment.NewLine +
                          "                       sum(isnull(b.CarCount, 0)) CarCount, sum(isnull(a.CBus1Km, 0)) + sum(isnull(a.CBus2Km, 0)) + sum(isnull(a.CBus4Km, 0)) as Bus_KM, sum(isnull(a.CRentTraKm, 0)) + sum(isnull(a.CBus3Km, 0)) as TransKM, " + Environment.NewLine +
                          "                       sum(isnull(a.CRentBKm, 0)) as TourKM, sum(isnull(a.CBus5KM, 0)) as NoneKM, " + Environment.NewLine +
                          "                       isnull((select sum(isnull(Amount, 0)) TAMT from SheetB " + Environment.NewLine +
                          "                                where SheetNo in (select SheetNo  from SheetA where SheetClass = '1' and SheetType = 'E' and SheetDiff = 'O' and Day between '" + vBuDate_S + "' and '" + vBuDate_E + "') " + Environment.NewLine +
                          "                                  and [No] = '7100120000' and DepNo = a.DepNo and Car_ID = a.Car_ID), 0) as DieselAMT, " + Environment.NewLine +
                          "                       isnull((select sum(isnull(Amount, 0)) TAMT from SheetB " + Environment.NewLine +
                          "                                where SheetNo in (select SheetNo  from SheetA where SheetClass = '1' and SheetType = 'E' and SheetDiff = 'O' and Day between '" + vBuDate_S + "' and '" + vBuDate_E + "') " + Environment.NewLine +
                          "                                  and [No] = '7100130000' and DepNo = a.DepNo  and Car_ID = a.Car_ID), 0) as EngineOilAMT, " + Environment.NewLine +
                          "                       isnull((select sum(Month_Dep) from Asset where AssetNo = a.Car_ID and (select CompanyNo from Car_InfoA where Car_ID = a.Car_ID) = a.DepNo), 0) as DepAMT, " + Environment.NewLine +
                          "                       isnull((select sum(isnull(Amount, 0)) TAMT from SheetB " + Environment.NewLine +
                          "                                where SheetNo in (select SheetNo  from SheetA where SheetClass = '1' and SheetType = 'E' and SheetDiff = 'M' and Day between '" + vBuDate_S + "' and '" + vBuDate_E + "') " + Environment.NewLine +
                          "                                  and DepNo = a.DepNo and Car_ID = a.Car_ID), 0) + " + Environment.NewLine +
                          "                       isnull((select sum(isnull(Amount, 0)) TAMT from FixoutB " + Environment.NewLine +
                          "                                where OutNo in (select OutNo  from FixOutA where AccountDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' and Mode <> '2') " + Environment.NewLine +
                          "                                  and DepNo = a.DepNo and Car_ID = a.Car_ID), 0) as FixAMT, " + Environment.NewLine +
                          "                       isnull((select sum(isnull(Amount, 0)) TAMT from SheetB " + Environment.NewLine +
                          "                                where SheetNo in (select SheetNo  from SheetA where SheetClass = '1' and SheetType = 'E' and SheetDiff = 'T' and Day between '" + vBuDate_S + "' and '" + vBuDate_E + "') " + Environment.NewLine +
                          "                                  and DepNo = a.DepNo and Car_ID = a.Car_ID), 0) + " + Environment.NewLine +
                          "                       isnull((select sum(isnull(Amount, 0)) TAMT from FixoutB " + Environment.NewLine +
                          "                                where OutNo in (select OutNo  from FixOutA where AccountDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' and Mode = '2') " + Environment.NewLine +
                          "                                  and DepNo = a.DepNo and Car_ID = a.Car_ID), 0) as TireAMT, " + Environment.NewLine +
                          "                       isnull((select sum(isnull(Debit, 0) - isnull(Credit, 0)) from Account " + Environment.NewLine +
                          "                                where left([Subject],4) in ('5183','5510','5520') and right([Subject], 2) = a.DepNo and [Type] <> '6' " + Environment.NewLine +
                          "                                  and isnull(Memo2, '') = a.Car_ID and Rec_Date between '" + vBuDate_S + "' and '" + vBuDate_E + "'), 0) LicenceAMT " + Environment.NewLine +
                          "                  from RsCarMonth a left join CarCount b on b.DepNo = a.DepNo and b.CalYM = '" + vBuDate_S + "' " + Environment.NewLine +
                          "                 where a.BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine +
                          //2023.12.27 增加從憑單判斷車號有沒有跑遊覽車趟次
                          //"                   and a.Car_ID in (select Car_ID from Car_InfoA where Point = 'C1' and Tran_Type = '1') " + Environment.NewLine +
                          "                   and (a.Car_ID in (select Car_ID from Car_InfoA where Point = 'C1' and Tran_Type = '1') " + Environment.NewLine +
                          "                    or a.Car_ID in (select distinct Car_ID from RunSheetB where CarType = '8' and AssignNo in (select AssignNo from RunSheetA where BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "') )) " + Environment.NewLine +
                          "                 group by a.DepNo, a.Car_ID, a.Car_No " + Environment.NewLine +
                          "               ) z " + Environment.NewLine +
                          "         group by z.DepNo, z.CarCOunt, z.TourCount " + Environment.NewLine +
                          "       ) x " + Environment.NewLine +
                          " order by x.DepNo";
                //宣告 SQL 連接器 
                SqlConnection connTemp = new SqlConnection(vConnStr);
                //宣告並初始化 SQL 資料集
                SqlDataAdapter daTemp = new SqlDataAdapter(vSQLStr, connTemp);
                //開啟資料庫連結
                connTemp.Open();
                //宣告暫存用的資料集
                DataSet dsTemp = new DataSet();
                //將 SQL 取回的資料填入暫存資料集
                daTemp.Fill(dsTemp, "TourCarIncome");

                string vCompanyNo = "";
                string vColumnsName = "";
                string vTempValue = "";

                if (dsTemp.Tables["TourCarIncome"].Rows.Count > 0) //如果暫存資料集有資料才進行後面的動作
                {
                    //每單位的個別資料計算寫入
                    for (int i = 0; i < dsTemp.Tables["TourCarIncome"].Rows.Count; i++)
                    {
                        vCompanyNo = dsTemp.Tables["TourCarIncome"].Rows[i]["DepNo"].ToString().Trim();
                        //取回各站賠償費
                        dsTemp.Tables["TourCarIncome"].Rows[i]["CompensateLoss"] = (vCompanyNo == "11") ? eCompensateLoss11.Text.Trim() :
                                                                                   (vCompanyNo == "13") ? eCompensateLoss13.Text.Trim() :
                                                                                   (vCompanyNo == "15") ? eCompensateLoss15.Text.Trim() :
                                                                                   (vCompanyNo == "16") ? eCompensateLoss16.Text.Trim() :
                                                                                   (vCompanyNo == "17") ? eCompensateLoss17.Text.Trim() :
                                                                                   (vCompanyNo == "18") ? eCompensateLoss18.Text.Trim() :
                                                                                   (vCompanyNo == "19") ? eCompensateLoss19.Text.Trim() :
                                                                                   (vCompanyNo == "20") ? eCompensateLoss20.Text.Trim() :
                                                                                   (vCompanyNo == "22") ? eCompensateLoss22.Text.Trim() :
                                                                                   (vCompanyNo == "23") ? eCompensateLoss23.Text.Trim() :
                                                                                   (vCompanyNo == "25") ? eCompensateLoss25.Text.Trim() : "0";
                        //加計專車收入
                        double vTempFee = double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TranIncome"].ToString().Trim());
                        string vTempFeeStr = (vCompanyNo == "11") ? eTranIncome11.Text.Trim() :
                                             (vCompanyNo == "13") ? eTranIncome13.Text.Trim() :
                                             (vCompanyNo == "15") ? eTranIncome15.Text.Trim() :
                                             (vCompanyNo == "16") ? eTranIncome16.Text.Trim() :
                                             (vCompanyNo == "17") ? eTranIncome17.Text.Trim() :
                                             (vCompanyNo == "18") ? eTranIncome18.Text.Trim() :
                                             (vCompanyNo == "19") ? eTranIncome19.Text.Trim() :
                                             (vCompanyNo == "20") ? eTranIncome20.Text.Trim() :
                                             (vCompanyNo == "22") ? eTranIncome22.Text.Trim() :
                                             (vCompanyNo == "23") ? eTranIncome23.Text.Trim() :
                                             (vCompanyNo == "25") ? eTranIncome25.Text.Trim() : "0.0";
                        dsTemp.Tables["TourCarIncome"].Rows[i]["TranIncome"] = (vTempFee + double.Parse(vTempFeeStr)).ToString().Trim();
                        //計算營業收入合計 TotalIncome (BusIncome + TranIncome + TourIncome)
                        vTempFee = double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["BusIncome"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TranIncome"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TourIncome"].ToString().Trim());
                        dsTemp.Tables["TourCarIncome"].Rows[i]["TotalIncome"] = vTempFee.ToString().Trim();
                        //計算修車員工費用 FixEmp
                        vSelectStr_Temp = "select cast(sum(SubjectAMT+OtherAMT+ShareAMT)  / " +
                                          "       (select CarCount from CarCount where CarCount.DepNoYM = a.IncomeYM + a.DepNo) * " +
                                          "       (select count(Car_ID) from Car_InfoA where Point = 'C1' and Tran_Type = '1' and Car_InfoA.CompanyNo = a.DepNo) as decimal(10, 0)) TAMT from DepMonthIncome a " +
                                          " where ChartBarCode = 'COS 65' and IncomeYM = '" + vYM_C + "' and DepNo = '" + vCompanyNo + "' " +
                                          " group by IncomeYM, DepNo";
                        vTempValue = PF.GetValue(vConnStr, vSelectStr_Temp, "TAMT");
                        dsTemp.Tables["TourCarIncome"].Rows[i]["FixEmp"] = (vTempValue.Trim() != "") ? vTempValue.Trim() : "0";
                        //計算支出合計 TotalOutpay (DieselAMT + EngineOilAMT + OtherOil + FixAMT + TireAMT + DepAMT + CarInst + CompensateLoss + LicenceAMT + DriverPayment + 
                        //                          BoundsFee + RetirePay + LiFee + HiFee + FixEmp + ClothCost)
                        vTempFee = double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["DieselAMT"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["EngineOilAMT"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["OtherOil"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["FixAMT"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TireAMT"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["DepAMT"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["CarInst"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["CompensateLoss"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["LicenceAMT"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["DriverPayment"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["BoundsFee"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["RetirePay"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["LiFee"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["HiFee"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["FixEmp"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["ClothCost"].ToString().Trim());
                        dsTemp.Tables["TourCarIncome"].Rows[i]["TotalOutpay"] = vTempFee.ToString().Trim();
                        //計算本期損益 TotalInOut (TotalIncome - TotalOutpay)
                        vTempFee = double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TotalIncome"].ToString().Trim()) -
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TotalOutpay"].ToString().Trim());
                        dsTemp.Tables["TourCarIncome"].Rows[i]["TotalInOut"] = vTempFee.ToString().Trim();
                        //計算里程合計 TotalKMs (Bus_KM + TransKM + TourKM + NoneKM)
                        vTempFee = double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["Bus_KM"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TransKM"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TourKM"].ToString().Trim()) +
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["NoneKM"].ToString().Trim());
                        dsTemp.Tables["TourCarIncome"].Rows[i]["TotalKMs"] = vTempFee.ToString().Trim();
                        //計算每車收入 CarIncome (TotalIncome / TourCount)
                        vTempFee = Math.Round((double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TotalIncome"].ToString().Trim()) /
                                               double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TourCount"].ToString().Trim())), MidpointRounding.AwayFromZero);
                        dsTemp.Tables["TourCarIncome"].Rows[i]["CarIncome"] = vTempFee.ToString().Trim();
                        //計算每車支出 CarOutpay (TotalOutpay / TourCount)
                        vTempFee = Math.Round((double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TotalOutpay"].ToString().Trim()) /
                                               double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TourCount"].ToString().Trim())), MidpointRounding.AwayFromZero);
                        dsTemp.Tables["TourCarIncome"].Rows[i]["CarOutpay"] = vTempFee.ToString().Trim();
                        //計算每車盈餘 CarInOut (CarIncome - CarOutpay)
                        vTempFee = double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["CarIncome"].ToString().Trim()) -
                                   double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["CarOutpay"].ToString().Trim());
                        dsTemp.Tables["TourCarIncome"].Rows[i]["CarInOut"] = vTempFee.ToString().Trim();
                    }
                    //計算合計
                    for (int i = 0; i < dsTemp.Tables["TourCarIncome"].Rows.Count; i++)
                    {
                        vTourIncome = vTourIncome + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TourIncome"].ToString().Trim()); //遊覽車收入
                        vTranIncome = vTranIncome + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TranIncome"].ToString().Trim()); //交通車收入
                        vBusIncome = vBusIncome + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["BusIncome"].ToString().Trim()); //班公車收入
                        vTotalIncome = vTotalIncome + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TotalIncome"].ToString().Trim()); //營業收入合計
                        vDieselAMT = vDieselAMT + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["DieselAMT"].ToString().Trim()); //柴油費
                        vEngineOilAMT = vEngineOilAMT + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["EngineOilAMT"].ToString().Trim()); //機油費
                        vOtherOil = vOtherOil + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["OtherOil"].ToString().Trim()); //其他油料費
                        vFixAMT = vFixAMT + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["FixAMT"].ToString().Trim()); //修車材料費
                        vTireAMT = vTireAMT + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TireAMT"].ToString().Trim()); //輪胎費
                        vDepAMT = vDepAMT + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["DepAMT"].ToString().Trim()); //車輛折舊
                        vCarInst = vCarInst + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["CarInst"].ToString().Trim()); //車輛保險費
                        vCompensateLoss = vCompensateLoss + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["CompensateLoss"].ToString().Trim()); //車輛賠償費
                        vLicenceAMT = vLicenceAMT + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["LicenceAMT"].ToString().Trim()); //證照費
                        vDriverPayment = vDriverPayment + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["DriverPayment"].ToString().Trim()); //駕駛員
                        vBoundsFee = vBoundsFee + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["BoundsFee"].ToString().Trim()); //獎金提列
                        vRetirePay = vRetirePay + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["RetirePay"].ToString().Trim()); //退休金
                        vLiFee = vLiFee + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["LiFee"].ToString().Trim()); //勞保費
                        vHiFee = vHiFee + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["HiFee"].ToString().Trim()); //健保費
                        vFixEmp = vFixEmp + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["FixEmp"].ToString().Trim()); //修車員工
                        vClothCost = vClothCost + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["ClothCost"].ToString().Trim()); //服裝費用
                        vTotalOutpay = vTotalOutpay + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TotalOutpay"].ToString().Trim()); //支出合計
                        vTotalInOut = vTotalInOut + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TotalInOut"].ToString().Trim()); //本期損益
                        vTourEmp = vTourEmp + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TourEmp"].ToString().Trim()); //駕駛員人數基準數
                        vTourCount = vTourCount + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TourCount"].ToString().Trim()); //使用車輛數基準數
                        vCarCount = vCarCount + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["CarCount"].ToString().Trim()); //站實際車輛數
                        vTourKM = vTourKM + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TourKM"].ToString().Trim()); //遊覽車里程
                        vTransKM = vTransKM + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TransKM"].ToString().Trim()); //交通車里程
                        vBus_KM = vBus_KM + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["Bus_KM"].ToString().Trim()); //班公車里程
                        vNoneKM = vNoneKM + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["NoneKM"].ToString().Trim()); //非營運里程
                        vTotalKMs = vTotalKMs + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["TotalKMs"].ToString().Trim()); //里程合計
                        //vCarIncome = vCarIncome + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["CarIncome"].ToString().Trim()); //每車收入
                        //vCarOutpay = vCarOutpay + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["CarOutpay"].ToString().Trim()); //每車支出
                        //vCarInOut = vCarInOut + double.Parse(dsTemp.Tables["TourCarIncome"].Rows[i]["CarInOut"].ToString().Trim()); //每車盈餘
                    }
                    DataRow drNewRow = dsTemp.Tables["TourCarIncome"].NewRow();
                    drNewRow["DepNo"] = "99";
                    drNewRow["TourIncome"] = vTourIncome.ToString(); //遊覽車收入
                    drNewRow["TranIncome"] = vTranIncome.ToString(); //交通車收入
                    drNewRow["BusIncome"] = vBusIncome.ToString(); //班公車收入
                    drNewRow["TotalIncome"] = vTotalIncome.ToString(); //營業收入合計
                    drNewRow["DieselAMT"] = vDieselAMT.ToString(); //柴油費
                    drNewRow["EngineOilAMT"] = vEngineOilAMT.ToString(); //機油費
                    drNewRow["OtherOil"] = vOtherOil.ToString(); //其他油料費
                    drNewRow["FixAMT"] = vFixAMT.ToString(); //修車材料費
                    drNewRow["TireAMT"] = vTireAMT.ToString(); //輪胎費
                    drNewRow["DepAMT"] = vDepAMT.ToString(); //車輛折舊
                    drNewRow["CarInst"] = vCarInst.ToString(); //車輛保險費
                    drNewRow["CompensateLoss"] = vCompensateLoss.ToString(); //車輛賠償費
                    drNewRow["LicenceAMT"] = vLicenceAMT.ToString(); //證照費
                    drNewRow["DriverPayment"] = vDriverPayment.ToString(); //駕駛員
                    drNewRow["BoundsFee"] = vBoundsFee.ToString(); //獎金提列
                    drNewRow["RetirePay"] = vRetirePay.ToString(); //退休金
                    drNewRow["LiFee"] = vLiFee.ToString(); //勞保費
                    drNewRow["HiFee"] = vHiFee.ToString(); //健保費
                    drNewRow["FixEmp"] = vFixEmp.ToString(); //修車員工
                    drNewRow["ClothCost"] = vClothCost.ToString(); //服裝費用
                    drNewRow["TotalOutpay"] = vTotalOutpay.ToString(); //支出合計
                    drNewRow["TotalInOut"] = vTotalInOut.ToString(); //本期損益
                    drNewRow["TourEmp"] = vTourEmp.ToString(); //駕駛員人數基準數
                    drNewRow["TourCount"] = vTourCount.ToString(); //使用車輛數基準數
                    drNewRow["CarCount"] = vCarCount.ToString(); //站實際車輛數
                    drNewRow["TourKM"] = vTourKM.ToString(); //遊覽車里程
                    drNewRow["TransKM"] = vTransKM.ToString(); //交通車里程
                    drNewRow["Bus_KM"] = vBus_KM.ToString(); //班公車里程
                    drNewRow["NoneKM"] = vNoneKM.ToString(); //非營運里程
                    drNewRow["TotalKMs"] = vTotalKMs.ToString(); //里程合計
                    //drNewRow["CarIncome"] = vCarIncome.ToString(); //每車收入 CarIncome (TotalIncome / TourCount)
                    drNewRow["CarIncome"] = Math.Round((vTotalIncome / vTourCount), MidpointRounding.AwayFromZero).ToString();
                    //drNewRow["CarOutpay"] = vCarOutpay.ToString(); //每車支出 CarOutpay (TotalOutpay / TourCount)
                    drNewRow["CarOutpay"] = Math.Round((vTotalOutpay / vTourCount), MidpointRounding.AwayFromZero).ToString();
                    //drNewRow["CarInOut"] = vCarInOut.ToString(); //每車盈餘
                    drNewRow["CarInOut"] = (Math.Round((vTotalIncome / vTourCount), MidpointRounding.AwayFromZero) - Math.Round((vTotalOutpay / vTourCount), MidpointRounding.AwayFromZero)).ToString();
                    drNewRow["CarYearInsu"] = "0"; //每車年強制險保費
                    drNewRow["CarThirdInsu"] = "0"; //每車年第三責任險
                    drNewRow["DriverLi"] = "0"; //駕駛員每人每月勞保費
                    drNewRow["DriverHi"] = "0"; //駕駛員每人每月健保費
                    drNewRow["DriverBounds"] = "0"; //駕駛員獎金每人每月
                    dsTemp.Tables["TourCarIncome"].Rows.Add(drNewRow);
                    //行列互換
                    //建立資料來源的結構
                    tbResult.Columns.Add("CalYM_C");
                    tbResult.Columns.Add("Items");
                    tbResult.Columns.Add("Items_C");
                    tbResult.Columns.Add("金額_11");
                    tbResult.Columns.Add("金額_13");
                    tbResult.Columns.Add("金額_15");
                    tbResult.Columns.Add("金額_16");
                    tbResult.Columns.Add("金額_17");
                    tbResult.Columns.Add("金額_18");
                    tbResult.Columns.Add("金額_19");
                    tbResult.Columns.Add("金額_20");
                    tbResult.Columns.Add("金額_22");
                    tbResult.Columns.Add("金額_23");
                    tbResult.Columns.Add("金額_25");
                    tbResult.Columns.Add("金額_99");
                    for (int i = 1; i < dsTemp.Tables["TourCarIncome"].Columns.Count; i++)
                    {
                        DataRow drTemp = tbResult.NewRow();
                        drTemp["CalYM_C"] = vYM_C;
                        drTemp["Items"] = dsTemp.Tables["TourCarIncome"].Columns[i].Caption;
                        drTemp["Items_C"] = (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "TourIncome") ? "遊覽車收入" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "TranIncome") ? "交通車收入" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "BusIncome") ? "班公車收入" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "TotalIncome") ? "營業收入合計" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "DieselAMT") ? "柴油費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "EngineOilAMT") ? "機油費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "OtherOil") ? "其他油料費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "FixAMT") ? "修車材料費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "TireAMT") ? "輪胎費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "DepAMT") ? "車輛折舊" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "CarInst") ? "車輛保險費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "CompensateLoss") ? "賠償費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "LicenceAMT") ? "憑照費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "DriverPayment") ? "駕駛員" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "BoundsFee") ? "獎金提列" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "RetirePay") ? "退休金" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "LiFee") ? "勞保費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "HiFee") ? "健保費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "FixEmp") ? "修車員工" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "ClothCost") ? "服裝費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "TotalOutpay") ? "支出合計" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "TotalInOut") ? "本期損益" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "TourEmp") ? "駕駛員人數基準數" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "TourCount") ? "使用車輛數基準數" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "CarCount") ? "站實際車輛數" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "TourKM") ? "遊覽車里程" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "TransKM") ? "交通車里程" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "Bus_KM") ? "班公車里程" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "NoneKM") ? "非營運里程" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "TotalKMs") ? "里程合計" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "CarIncome") ? "每車收入" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "CarOutpay") ? "每車支出" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "CarInOut") ? "每車盈餘" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "CarYearInsu") ? "每車年強制保險費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "CarThirdInsu") ? "每車年第三責任險" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "DriverLi") ? "駕駛員每人月勞保費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "DriverHi") ? "駕駛員每人月健保費" :
                                            (dsTemp.Tables["TourCarIncome"].Columns[i].Caption == "DriverBounds") ? "駕駛員獎金每人每月" : "";
                        for (int j = 0; j < dsTemp.Tables["TourCarIncome"].Rows.Count; j++)
                        {
                            vCompanyNo = dsTemp.Tables["TourCarIncome"].Rows[j]["DepNo"].ToString().Trim();
                            vColumnsName = "金額_" + vCompanyNo;
                            vRowsName = (dsTemp.Tables["TourCarIncome"].Rows[j][i].ToString().Trim() != "") ? dsTemp.Tables["TourCarIncome"].Rows[j][i].ToString().Trim() : "0";
                            vRowsName = (vRowsName.Trim() == "") ? "0" : vRowsName.Trim();
                            drTemp[vColumnsName] = vRowsName;
                        }
                        tbResult.Rows.Add(drTemp);
                    }

                    tbResult.Columns.Add("公里收支_11");
                    tbResult.Columns.Add("公里收支_13");
                    tbResult.Columns.Add("公里收支_15");
                    tbResult.Columns.Add("公里收支_16");
                    tbResult.Columns.Add("公里收支_17");
                    tbResult.Columns.Add("公里收支_18");
                    tbResult.Columns.Add("公里收支_19");
                    tbResult.Columns.Add("公里收支_20");
                    tbResult.Columns.Add("公里收支_22");
                    tbResult.Columns.Add("公里收支_23");
                    tbResult.Columns.Add("公里收支_25");
                    tbResult.Columns.Add("公里收支_99");
                    for (int i = 0; i < tbResult.Rows.Count; i++)
                    {
                        if (ItemList.Any(tbResult.Rows[i]["Items"].ToString().Trim().Contains))
                        {
                            if (tbResult.Rows[i]["Items"].ToString().Trim() == "TourIncome") //遊覽車收入用遊覽車里程計算
                            {
                                for (int k = 0; k < DepList.Count; k++)
                                {
                                    vCalDepNo = DepList[k].ToString().Trim();
                                    vInOutFieldName = "公里收支_" + vCalDepNo;
                                    vCashFieldName = "金額_" + vCalDepNo;
                                    vSelectStr_Temp = "select isnull(cast(sum(isnull(CRentBKm, 0)) as decimal(10,0)), 0) TotalKM from RSCarMonth " +
                                                      " where Car_ID in (select Car_ID from Car_InfoA where Point = 'C1' and Tran_Type = '1') " +
                                                      "   and DepNo = '" + vCalDepNo + "' " +
                                                      "   and BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "'";
                                    tbResult.Rows[i][vInOutFieldName] = ((tbResult.Rows[i][vCashFieldName].ToString().Trim() != "") && (tbResult.Rows[i][vCashFieldName].ToString().Trim() != "0")) ?
                                                                        Math.Round((double.Parse(tbResult.Rows[i][vCashFieldName].ToString().Trim()) /
                                                                                    double.Parse(PF.GetValue(vConnStr, vSelectStr_Temp, "TotalKM").Trim())),
                                                                                   2,
                                                                                   MidpointRounding.AwayFromZero).ToString() : "0";
                                }
                                vSelectStr_Temp = "select isnull(cast(sum(isnull(CRentBKm, 0)) as decimal(10,0)), 0) TotalKM from RSCarMonth " +
                                                  " where Car_ID in (select Car_ID from Car_InfoA where Point = 'C1' and Tran_Type = '1') " +
                                                  "   and BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "'";
                                tbResult.Rows[i]["公里收支_99"] = ((tbResult.Rows[i]["金額_99"].ToString().Trim() != "") && (tbResult.Rows[i]["金額_99"].ToString().Trim() != "0")) ?
                                                                  Math.Round((double.Parse(tbResult.Rows[i]["金額_99"].ToString().Trim()) /
                                                                              double.Parse(PF.GetValue(vConnStr, vSelectStr_Temp, "TotalKM").Trim())),
                                                                             2,
                                                                             MidpointRounding.AwayFromZero).ToString() : "0";
                            }
                            else if (tbResult.Rows[i]["Items"].ToString().Trim() == "TranIncome") //交通車收入用交通車里程計算
                            {
                                for (int k = 0; k < DepList.Count; k++)
                                {
                                    vCalDepNo = DepList[k].ToString().Trim();
                                    vInOutFieldName = "公里收支_" + vCalDepNo;
                                    vCashFieldName = "金額_" + vCalDepNo;
                                    vSelectStr_Temp = "select sum(isnull(CRentTraKm, 0)) + sum(isnull(CBus3Km, 0)) TotalKM from RSCarMonth " +
                                                      " where Car_ID in (select Car_ID from Car_InfoA where Point = 'C1' and Tran_Type = '1') " +
                                                      "   and DepNo = '" + vCalDepNo + "' " +
                                                      "   and BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "'";
                                    tbResult.Rows[i][vInOutFieldName] = ((tbResult.Rows[i][vCashFieldName].ToString().Trim() != "") && (tbResult.Rows[i][vCashFieldName].ToString().Trim() != "0")) ?
                                                                        Math.Round((double.Parse(tbResult.Rows[i][vCashFieldName].ToString().Trim()) /
                                                                                    double.Parse(PF.GetValue(vConnStr, vSelectStr_Temp, "TotalKM").Trim())),
                                                                                   2,
                                                                                   MidpointRounding.AwayFromZero).ToString() : "0";
                                }
                                vSelectStr_Temp = "select sum(isnull(CRentTraKm, 0)) + sum(isnull(CBus3Km, 0)) TotalKM from RSCarMonth " +
                                                  " where Car_ID in (select Car_ID from Car_InfoA where Point = 'C1' and Tran_Type = '1') " +
                                                  "   and BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "'";
                                tbResult.Rows[i]["公里收支_99"] = ((tbResult.Rows[i]["金額_99"].ToString().Trim() != "") && (tbResult.Rows[i]["金額_99"].ToString().Trim() != "0")) ?
                                                                  Math.Round((double.Parse(tbResult.Rows[i]["金額_99"].ToString().Trim()) /
                                                                              double.Parse(PF.GetValue(vConnStr, vSelectStr_Temp, "TotalKM").Trim())),
                                                                             2,
                                                                             MidpointRounding.AwayFromZero).ToString() : "0";
                            }
                            else if (tbResult.Rows[i]["Items"].ToString().Trim() == "BusIncome") //班公車收入用班公車里程加非營運里程計算
                            {
                                for (int k = 0; k < DepList.Count; k++)
                                {
                                    vCalDepNo = DepList[k].ToString().Trim();
                                    vInOutFieldName = "公里收支_" + vCalDepNo;
                                    vCashFieldName = "金額_" + vCalDepNo;
                                    vSelectStr_Temp = "select isnull(cast(sum(isnull(CBus1Km, 0)) + sum(isnull(CBus2Km, 0)) + sum(isnull(CBus4Km, 0)) + sum(isnull(CBus5KM, 0)) as decimal(10,0)), 0) TotalKM from RSCarMonth " +
                                                      " where Car_ID in (select Car_ID from Car_InfoA where Point = 'C1' and Tran_Type = '1') " +
                                                      "   and DepNo = '" + vCalDepNo + "' " +
                                                      "   and BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "'";
                                    tbResult.Rows[i][vInOutFieldName] = ((tbResult.Rows[i][vCashFieldName].ToString().Trim() != "") && (tbResult.Rows[i][vCashFieldName].ToString().Trim() != "0")) ?
                                                                        Math.Round((double.Parse(tbResult.Rows[i][vCashFieldName].ToString().Trim()) /
                                                                                    double.Parse(PF.GetValue(vConnStr, vSelectStr_Temp, "TotalKM").Trim())),
                                                                                   2,
                                                                                   MidpointRounding.AwayFromZero).ToString() : "0";
                                }
                                vSelectStr_Temp = "select isnull(cast(sum(isnull(CBus1Km, 0)) + sum(isnull(CBus2Km, 0)) + sum(isnull(CBus4Km, 0)) + sum(isnull(CBus5KM, 0)) as decimal(10,0)), 0) TotalKM from RSCarMonth " +
                                                  " where Car_ID in (select Car_ID from Car_InfoA where Point = 'C1' and Tran_Type = '1') " +
                                                  "   and BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "'";
                                tbResult.Rows[i]["公里收支_99"] = ((tbResult.Rows[i]["金額_99"].ToString().Trim() != "") && (tbResult.Rows[i]["金額_99"].ToString().Trim() != "0")) ?
                                                                  Math.Round((double.Parse(tbResult.Rows[i]["金額_99"].ToString().Trim()) /
                                                                              double.Parse(PF.GetValue(vConnStr, vSelectStr_Temp, "TotalKM").Trim())),
                                                                             2,
                                                                             MidpointRounding.AwayFromZero).ToString() : "0";
                            }
                            else //其他的項目都用里程合計計算
                            {
                                for (int k = 0; k < DepList.Count; k++)
                                {
                                    vCalDepNo = DepList[k].ToString().Trim();
                                    vInOutFieldName = "公里收支_" + vCalDepNo;
                                    vCashFieldName = "金額_" + vCalDepNo;
                                    vSelectStr_Temp = "select isnull(cast(sum(isnull(CRentBKm, 0)) + sum(isnull(CBus1Km, 0)) + sum(isnull(CBus2Km, 0)) + sum(isnull(CBus4Km, 0)) + sum(isnull(CBus5KM, 0)) + sum(isnull(CRentTraKm, 0)) + sum(isnull(CBus3Km, 0)) as decimal(10,0)), 0) TotalKM from RSCarMonth " +
                                                      " where Car_ID in (select Car_ID from Car_InfoA where Point = 'C1' and Tran_Type = '1') " +
                                                      "   and DepNo = '" + vCalDepNo + "' " +
                                                      "   and BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "'";
                                    tbResult.Rows[i][vInOutFieldName] = ((tbResult.Rows[i][vCashFieldName].ToString().Trim() != "") && (tbResult.Rows[i][vCashFieldName].ToString().Trim() != "0")) ?
                                                                        Math.Round((double.Parse(tbResult.Rows[i][vCashFieldName].ToString().Trim()) /
                                                                                    double.Parse(PF.GetValue(vConnStr, vSelectStr_Temp, "TotalKM").Trim())),
                                                                                   2,
                                                                                   MidpointRounding.AwayFromZero).ToString() : "0";
                                }
                                vSelectStr_Temp = "select isnull(cast(sum(isnull(CRentBKm, 0)) + sum(isnull(CBus1Km, 0)) + sum(isnull(CBus2Km, 0)) + sum(isnull(CBus4Km, 0)) + sum(isnull(CBus5KM, 0)) + sum(isnull(CRentTraKm, 0)) + sum(isnull(CBus3Km, 0)) as decimal(10,0)), 0) TotalKM from RSCarMonth " +
                                                  " where Car_ID in (select Car_ID from Car_InfoA where Point = 'C1' and Tran_Type = '1') " +
                                                  "   and BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "'";
                                tbResult.Rows[i]["公里收支_99"] = ((tbResult.Rows[i]["金額_99"].ToString().Trim() != "") && (tbResult.Rows[i]["金額_99"].ToString().Trim() != "0")) ?
                                                                  Math.Round((double.Parse(tbResult.Rows[i]["金額_99"].ToString().Trim()) /
                                                                              double.Parse(PF.GetValue(vConnStr, vSelectStr_Temp, "TotalKM").Trim())),
                                                                             2,
                                                                             MidpointRounding.AwayFromZero).ToString() : "0";
                            }
                        }
                    }
                    if (tbResult.Rows.Count > 0)
                    {
                        //預覽報表
                        ReportDataSource rdsPrint = new ReportDataSource("TourCarIncome", tbResult);
                        rvPrint.LocalReport.DataSources.Clear(); //清除報表資料來源
                        rvPrint.LocalReport.ReportPath = @"Report\TourCarIncomeP.rdlc";
                        rvPrint.LocalReport.DataSources.Add(rdsPrint); //把自建的 DataTable 指定給報表
                        rvPrint.LocalReport.Refresh();
                        plPrint.Visible = true;
                        plInputData.Visible = false;

                        string vDriveYMStr = DriveYM.Text.Trim();
                        string vCarYearInsuranceStr = (CarYearInsurance.Text.Trim() != "") ? CarYearInsurance.Text.Trim() : "";
                        string vCarThirdInsuranceStr = (CarThirdInsurance.Text.Trim() != "") ? CarThirdInsurance.Text.Trim() : "";
                        string vDriverLiMonthStr = (DriverLiMonth.Text.Trim() != "") ? DriverLiMonth.Text.Trim() : "";
                        string vDriverHiMonthStr = (DriverHiMonth.Text.Trim() != "") ? DriverHiMonth.Text.Trim() : "";
                        string vDriverBoundsMonthStr = (DriverBoundsMonth.Text.Trim() != "") ? DriverBoundsMonth.Text.Trim() : "";
                        string vCompensateLoss11Str = (eCompensateLoss11.Text.Trim() != "") ? eCompensateLoss11.Text.Trim() : "";
                        string vCompensateLoss13Str = (eCompensateLoss13.Text.Trim() != "") ? eCompensateLoss13.Text.Trim() : "";
                        string vCompensateLoss15Str = (eCompensateLoss15.Text.Trim() != "") ? eCompensateLoss15.Text.Trim() : "";
                        string vCompensateLoss16Str = (eCompensateLoss16.Text.Trim() != "") ? eCompensateLoss16.Text.Trim() : "";
                        string vCompensateLoss17Str = (eCompensateLoss17.Text.Trim() != "") ? eCompensateLoss17.Text.Trim() : "";
                        string vCompensateLoss18Str = (eCompensateLoss18.Text.Trim() != "") ? eCompensateLoss18.Text.Trim() : "";
                        string vCompensateLoss19Str = (eCompensateLoss19.Text.Trim() != "") ? eCompensateLoss19.Text.Trim() : "";
                        string vCompensateLoss20Str = (eCompensateLoss20.Text.Trim() != "") ? eCompensateLoss20.Text.Trim() : "";
                        string vCompensateLoss22Str = (eCompensateLoss22.Text.Trim() != "") ? eCompensateLoss22.Text.Trim() : "";
                        string vCompensateLoss23Str = (eCompensateLoss23.Text.Trim() != "") ? eCompensateLoss23.Text.Trim() : "";
                        string vCompensateLoss25Str = (eCompensateLoss25.Text.Trim() != "") ? eCompensateLoss25.Text.Trim() : "";
                        string vTranIncome11Str = (eTranIncome11.Text.Trim() != "") ? eTranIncome11.Text.Trim() : "";
                        string vTranIncome13Str = (eTranIncome13.Text.Trim() != "") ? eTranIncome13.Text.Trim() : "";
                        string vTranIncome15Str = (eTranIncome15.Text.Trim() != "") ? eTranIncome15.Text.Trim() : "";
                        string vTranIncome16Str = (eTranIncome16.Text.Trim() != "") ? eTranIncome16.Text.Trim() : "";
                        string vTranIncome17Str = (eTranIncome17.Text.Trim() != "") ? eTranIncome17.Text.Trim() : "";
                        string vTranIncome18Str = (eTranIncome18.Text.Trim() != "") ? eTranIncome18.Text.Trim() : "";
                        string vTranIncome19Str = (eTranIncome19.Text.Trim() != "") ? eTranIncome19.Text.Trim() : "";
                        string vTranIncome20Str = (eTranIncome20.Text.Trim() != "") ? eTranIncome20.Text.Trim() : "";
                        string vTranIncome22Str = (eTranIncome22.Text.Trim() != "") ? eTranIncome22.Text.Trim() : "";
                        string vTranIncome23Str = (eTranIncome23.Text.Trim() != "") ? eTranIncome23.Text.Trim() : "";
                        string vTranIncome25Str = (eTranIncome25.Text.Trim() != "") ? eTranIncome25.Text.Trim() : "";
                        string vClothCostStr = (ClothCost.Text.Trim() != "") ? ClothCost.Text.Trim() : "";
                        string vRecordNote = "匯出檔案_各站遊覽車營收成本表" + Environment.NewLine +
                                             "TourCarIncome.aspx" + Environment.NewLine +
                                             "行車年月：" + vDriveYMStr + Environment.NewLine +
                                             "每車年強制險保費：" + vCarYearInsuranceStr + Environment.NewLine +
                                             "每車年第三責任險：" + vCarThirdInsuranceStr + Environment.NewLine +
                                             "駕駛員每人月勞保費：" + vDriverLiMonthStr + Environment.NewLine +
                                             "駕駛員每人月健保費：" + vDriverHiMonthStr + Environment.NewLine +
                                             "駕駛員獎金每人每月：" + vDriverBoundsMonthStr + Environment.NewLine +
                                             "賠償費：桃園站~" + vCompensateLoss11Str + Environment.NewLine +
                                             "        大溪站~" + vCompensateLoss13Str + Environment.NewLine +
                                             "        大園站~" + vCompensateLoss15Str + Environment.NewLine +
                                             "        觀音站~" + vCompensateLoss16Str + Environment.NewLine +
                                             "        龍潭站~" + vCompensateLoss17Str + Environment.NewLine +
                                             "        竹圍站~" + vCompensateLoss18Str + Environment.NewLine +
                                             "        三峽站~" + vCompensateLoss19Str + Environment.NewLine +
                                             "        新屋站~" + vCompensateLoss20Str + Environment.NewLine +
                                             "        桃園公車站~" + vCompensateLoss22Str + Environment.NewLine +
                                             "        中壢公車站~" + vCompensateLoss23Str + Environment.NewLine +
                                             "        中壢站公車~" + vCompensateLoss25Str + Environment.NewLine +
                                             "專車收入：桃園站~" + vTranIncome11Str + Environment.NewLine +
                                             "          大溪站~" + vTranIncome13Str + Environment.NewLine +
                                             "          大園站~" + vTranIncome15Str + Environment.NewLine +
                                             "          觀音站~" + vTranIncome16Str + Environment.NewLine +
                                             "          龍潭站~" + vTranIncome17Str + Environment.NewLine +
                                             "          竹圍站~" + vTranIncome18Str + Environment.NewLine +
                                             "          三峽站~" + vTranIncome19Str + Environment.NewLine +
                                             "          新屋站~" + vTranIncome20Str + Environment.NewLine +
                                             "          桃園公車站~" + vTranIncome22Str + Environment.NewLine +
                                             "          中壢公車站~" + vTranIncome23Str + Environment.NewLine +
                                             "          中壢站公車~" + vTranIncome25Str + Environment.NewLine +
                                             "每人服裝費：" + vClothCostStr;
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('查無符合條件資料！')");
                    Response.Write("</" + "Script>");
                    DriveYM.Text = "";
                    DriveYM.Focus();
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('行車年月格式輸入有誤！')");
                Response.Write("</" + "Script>");
                DriveYM.Text = "";
                DriveYM.Focus();
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbClosePrint_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plInputData.Visible = true;
        }
    }
}