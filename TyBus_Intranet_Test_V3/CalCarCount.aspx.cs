using Amaterasu_Function;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class CalCarCount : Page
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
        private string vErrorMessage = "";
        private DataTable dtTarget = new DataTable();

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

                UnobtrusiveValidationMode = UnobtrusiveValidationMode.None;

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }

                    if (!IsPostBack)
                    {

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

        protected void bbCalData_Click(object sender, EventArgs e)
        {
            string vDepNo = "";
            string vSelectStr = "";
            string vTempSelectStr = "";
            double vTempFloat = 0.0;
            double vTempFloatA = 0.0;
            double vTempFloatB = 0.0;
            double vTempFloatT = 0.0;
            double vTempRuleKMS = 0.0;
            double vTempRuleKMS_T = 0.0;
            double vTempBusKMS = 0.0;
            double vTempBusKMS_T = 0.0;
            double vTempSpecKMS = 0.0;
            double vTempSpecKMS_T = 0.0;
            double vTempRentKMS = 0.0;
            double vTempRentKMS_T = 0.0;
            double vTempTransKMS = 0.0;
            double vTempTransKMS_T = 0.0;
            double vTempTourKMS = 0.0;
            double vTempTourKMS_T = 0.0;
            double vTempNoneBusiKMS = 0.0;
            double vTempNoneBusiKMS_T = 0.0;
            double vEmpCount = 0.0;
            double vEmpCount_M = 0.0;
            double vEmpCount_F = 0.0;
            double vEmpCount_D = 0.0;
            string vODYM = "";
            double vDriverOD = 0.0;
            int vTempINT = 0;
            bool vHasOldData = false;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if (eCalYM_Search.Text.Trim() != "")
            {
                //用 System.Linq 提供的函式 string.All() 來檢查字串是不是都是數字
                if ((eCalYM_Search.Text.Trim().Length == 5) && (eCalYM_Search.Text.Trim().All(char.IsDigit)))
                {
                    DateTime eCalYM_Temp = new DateTime(Int32.Parse(eCalYM_Search.Text.Trim().Substring(0, 3)) + 1911, Int32.Parse(eCalYM_Search.Text.Trim().Substring(3, 2)), 1);
                    vHasOldData = (Int32.Parse(PF.GetValue(vConnStr, "select count(DepNoYM) RCount from CarCount where CalYM = '" + eCalYM_Temp.ToString("yyyy/MM/dd") + "' ", "RCount")) > 0);
                    DateTime eCalYMLast_Temp = DateTime.Parse(PF.GetMonthLastDay(eCalYM_Temp, "B"), new CultureInfo("en-US"));
                    /* 2024.02.07 改從 Department 取回基本資料清單
                    vSelectStr = "select ('" + eCalYM_Search.Text.Trim() + "' + CompanyNo) as DepNoYM, cast('" + eCalYM_Temp.ToString("yyyy/MM/dd") + "' as DateTime) as CalYM, CompanyNo as DepNo, " + Environment.NewLine +
                                 "       cast(null as datetime) ModifyYM, cast(null as varchar)ModifyMan, cast(0 as float) CarCount, cast(0 as float) DriveRange, cast(null as varchar) Remark, " + Environment.NewLine +
                                 "       cast(0 as float) KMS_BUS, cast(0 as float) KMS_Rent, cast(0 as float) KMS_Tour, cast(0 as float) KMS_Trans, cast(0 as float) KMS_Spec, " + Environment.NewLine +
                                 "       cast(0 as float) RuleCount_OA, cast(0 as float) RuleCount_EA, cast(0 as float) RuleCount_OB, cast(0 as float) RuleCount_EB, cast(0 as float) RuleCount_OT, cast(0 as float) RuleCount_ET, " + Environment.NewLine +
                                 "       cast(0 as float) BusCount_OA, cast(0 as float) BusCount_EA, cast(0 as float) BusCount_OB, cast(0 as float) BusCount_EB, cast(0 as float) BusCount_OT, cast(0 as float) BusCount_ET, " + Environment.NewLine +
                                 "       cast(0 as float) TourCount_OA, cast(0 as float) TourCount_EA, cast(0 as float) TourCount_OB, cast(0 as float) TourCount_EB, cast(0 as float) TourCount_OT, cast(0 as float) TourCount_ET, " + Environment.NewLine +
                                 "       cast(0 as float) CarSum_OA, cast(0 as float) CarSum_EA, cast(0 as float) CarSum_OB, cast(0 as float) CarSum_EB, cast(0 as float) CarSum_OT, cast(0 as float) CarSum_ET, " + Environment.NewLine +
                                 "       cast(0 as float) RuleKMS_OA, cast(0 as float) RuleKMS_EA, cast(0 as float) RuleKMS_OB, cast(0 as float) RuleKMS_EB, cast(0 as float) RuleKMS_OT, cast(0 as float) RuleKMS_ET, cast(0 as float) KMS_Rule, " + Environment.NewLine +
                                 "       cast(0 as float) BusKMS_OA, cast(0 as float) BusKMS_EA, cast(0 as float) BusKMS_OB, cast(0 as float) BusKMS_EB, cast(0 as float) BusKMS_OT, cast(0 as float) BusKMS_ET, " + Environment.NewLine +
                                 "       cast(0 as float) SpecKMS_OA, cast(0 as float) SpecKMS_EA, cast(0 as float) SpecKMS_OB, cast(0 as float) SpecKMS_EB, cast(0 as float) SpecKMS_OT, cast(0 as float) SpecKMS_ET, " + Environment.NewLine +
                                 "       cast(0 as float) RentKMS_OA, cast(0 as float) RentKMS_EA, cast(0 as float) RentKMS_OB, cast(0 as float) RentKMS_EB, cast(0 as float) RentKMS_OT, cast(0 as float) RentKMS_ET, " + Environment.NewLine +
                                 "       cast(0 as float) TransKMS_OA, cast(0 as float) TransKMS_EA, cast(0 as float) TransKMS_OB, cast(0 as float) TransKMS_EB, cast(0 as float) TransKMS_OT, cast(0 as float) TransKMS_ET, " + Environment.NewLine +
                                 "       cast(0 as float) TourKMS_OA, cast(0 as float) TourKMS_EA, cast(0 as float) TourKMS_OB, cast(0 as float) TourKMS_EB, cast(0 as float) TourKMS_OT, cast(0 as float) TourKMS_ET, " + Environment.NewLine +
                                 "       cast(0 as float) NoneBusiKMS_OA, cast(0 as float) NoneBusiKMS_EA, cast(0 as float) NoneBusiKMS_OB, cast(0 as float) NoneBusiKMS_EB, cast(0 as float) NoneBusiKMS_OT, cast(0 as float) NoneBusiKMS_ET, cast(0 as float) KMS_NoneBusi, " + Environment.NewLine +
                                 "       cast(0 as float) TotalKMS_OA, cast(0 as float) TotalKMS_EA, cast(0 as float) TotalKMS_OB, cast(0 as float) TotalKMS_EB, cast(0 as float) TotalKMS_OT, cast(0 as float) TotalKMS_ET, " + Environment.NewLine +
                                 "       cast(0 as float) KMS_Total, cast(0 as float) HighWayCount_OA, cast(0 as float) HighWayCount_EA, cast(0 as float) HighWayCount_OB, cast(0 as float) HighWayCount_EB, cast(0 as float) HighWayCount_OT, cast(0 as float) HighWayCount_ET, " + Environment.NewLine +
                                 "       cast(0 as float) OtherCount_OA, cast(0 as float) OtherCount_EA, cast(0 as float) OtherCount_OB, cast(0 as float) OtherCount_EB, cast(0 as float) OtherCount_OT, cast(0 as float) OtherCount_ET, " + Environment.NewLine +
                                 "       cast(0 as float) ModifyKMS, cast(0 as float) SpecIncome, cast(0 as float) UsedCars, cast(0 as float) ElecCardCount, cast(0 as float) EmpCount, cast(0 as float) DriverOD, cast(0 as float) EmpCount_M, cast(0 as float) EmpCount_F, cast(0 as float) EmpCount_D, cast(0 as float) BreakDown " + Environment.NewLine +
                                 "  from Car_infoA group by CompanyNo "; */
                    vSelectStr = "select ('" + eCalYM_Search.Text.Trim() + "' + DepNo) as DepNoYM, cast('" + eCalYM_Temp.ToString("yyyy/MM/dd") + "' as DateTime) as CalYM, DepNo, " + Environment.NewLine +
                                 "       cast(null as datetime) ModifyYM, cast(null as varchar)ModifyMan, cast(0 as float) CarCount, cast(0 as float) DriveRange, cast(null as varchar) Remark, " + Environment.NewLine +
                                 "       cast(0 as float) KMS_BUS, cast(0 as float) KMS_Rent, cast(0 as float) KMS_Tour, cast(0 as float) KMS_Trans, cast(0 as float) KMS_Spec, " + Environment.NewLine +
                                 "       cast(0 as float) RuleCount_OA, cast(0 as float) RuleCount_EA, cast(0 as float) RuleCount_OB, cast(0 as float) RuleCount_EB, cast(0 as float) RuleCount_OT, cast(0 as float) RuleCount_ET, " + Environment.NewLine +
                                 "       cast(0 as float) BusCount_OA, cast(0 as float) BusCount_EA, cast(0 as float) BusCount_OB, cast(0 as float) BusCount_EB, cast(0 as float) BusCount_OT, cast(0 as float) BusCount_ET, " + Environment.NewLine +
                                 "       cast(0 as float) TourCount_OA, cast(0 as float) TourCount_EA, cast(0 as float) TourCount_OB, cast(0 as float) TourCount_EB, cast(0 as float) TourCount_OT, cast(0 as float) TourCount_ET, " + Environment.NewLine +
                                 "       cast(0 as float) CarSum_OA, cast(0 as float) CarSum_EA, cast(0 as float) CarSum_OB, cast(0 as float) CarSum_EB, cast(0 as float) CarSum_OT, cast(0 as float) CarSum_ET, " + Environment.NewLine +
                                 "       cast(0 as float) RuleKMS_OA, cast(0 as float) RuleKMS_EA, cast(0 as float) RuleKMS_OB, cast(0 as float) RuleKMS_EB, cast(0 as float) RuleKMS_OT, cast(0 as float) RuleKMS_ET, cast(0 as float) KMS_Rule, " + Environment.NewLine +
                                 "       cast(0 as float) BusKMS_OA, cast(0 as float) BusKMS_EA, cast(0 as float) BusKMS_OB, cast(0 as float) BusKMS_EB, cast(0 as float) BusKMS_OT, cast(0 as float) BusKMS_ET, " + Environment.NewLine +
                                 "       cast(0 as float) SpecKMS_OA, cast(0 as float) SpecKMS_EA, cast(0 as float) SpecKMS_OB, cast(0 as float) SpecKMS_EB, cast(0 as float) SpecKMS_OT, cast(0 as float) SpecKMS_ET, " + Environment.NewLine +
                                 "       cast(0 as float) RentKMS_OA, cast(0 as float) RentKMS_EA, cast(0 as float) RentKMS_OB, cast(0 as float) RentKMS_EB, cast(0 as float) RentKMS_OT, cast(0 as float) RentKMS_ET, " + Environment.NewLine +
                                 "       cast(0 as float) TransKMS_OA, cast(0 as float) TransKMS_EA, cast(0 as float) TransKMS_OB, cast(0 as float) TransKMS_EB, cast(0 as float) TransKMS_OT, cast(0 as float) TransKMS_ET, " + Environment.NewLine +
                                 "       cast(0 as float) TourKMS_OA, cast(0 as float) TourKMS_EA, cast(0 as float) TourKMS_OB, cast(0 as float) TourKMS_EB, cast(0 as float) TourKMS_OT, cast(0 as float) TourKMS_ET, " + Environment.NewLine +
                                 "       cast(0 as float) NoneBusiKMS_OA, cast(0 as float) NoneBusiKMS_EA, cast(0 as float) NoneBusiKMS_OB, cast(0 as float) NoneBusiKMS_EB, cast(0 as float) NoneBusiKMS_OT, cast(0 as float) NoneBusiKMS_ET, cast(0 as float) KMS_NoneBusi, " + Environment.NewLine +
                                 "       cast(0 as float) TotalKMS_OA, cast(0 as float) TotalKMS_EA, cast(0 as float) TotalKMS_OB, cast(0 as float) TotalKMS_EB, cast(0 as float) TotalKMS_OT, cast(0 as float) TotalKMS_ET, " + Environment.NewLine +
                                 "       cast(0 as float) KMS_Total, cast(0 as float) HighWayCount_OA, cast(0 as float) HighWayCount_EA, cast(0 as float) HighWayCount_OB, cast(0 as float) HighWayCount_EB, cast(0 as float) HighWayCount_OT, cast(0 as float) HighWayCount_ET, " + Environment.NewLine +
                                 "       cast(0 as float) OtherCount_OA, cast(0 as float) OtherCount_EA, cast(0 as float) OtherCount_OB, cast(0 as float) OtherCount_EB, cast(0 as float) OtherCount_OT, cast(0 as float) OtherCount_ET, " + Environment.NewLine +
                                 "       cast(0 as float) ModifyKMS, cast(0 as float) SpecIncome, cast(0 as float) UsedCars, cast(0 as float) ElecCardCount, cast(0 as float) EmpCount, cast(0 as float) DriverOD, cast(0 as float) EmpCount_M, cast(0 as float) EmpCount_F, cast(0 as float) EmpCount_D, cast(0 as float) BreakDown " + Environment.NewLine +
                                 "  from Department " + Environment.NewLine +
                                 " where InPayReport = 'Y' and IsPrint = 'Y' " + Environment.NewLine +
                                 " order by DepNo ";
                    using (SqlConnection connCarCount = new SqlConnection(vConnStr))
                    {
                        SqlDataAdapter daCarCount = new SqlDataAdapter(vSelectStr, connCarCount);
                        connCarCount.Open();
                        daCarCount.Fill(dtTarget);
                        if (dtTarget.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtTarget.Rows.Count; i++)
                            {
                                vDepNo = dtTarget.Rows[i]["DepNo"].ToString().Trim();
                                //先計算各類車種數量
                                vTempSelectStr = "select count(Car_ID) as CarCount, Point, Class, Exceptional " + Environment.NewLine +
                                                 "  from Car_InfoA " + Environment.NewLine +
                                                 " where TRAN_TYPE in ('1', '2') " + Environment.NewLine +
                                                 "   and CompanyNo = '" + vDepNo + "' " + Environment.NewLine +
                                                 " group by CompanyNo, Point, Class, Exceptional";
                                using (SqlConnection connCalCount = new SqlConnection(vConnStr))
                                {
                                    SqlCommand cmdCalCount = new SqlCommand(vTempSelectStr, connCalCount);
                                    connCalCount.Open();
                                    SqlDataReader drCalCount = cmdCalCount.ExecuteReader();
                                    while (drCalCount.Read())
                                    {
                                        vTempFloat = double.Parse(drCalCount["CarCount"].ToString().Trim());
                                        if (drCalCount["Point"].ToString().Trim().ToUpper() == "A1") //國道車
                                        {
                                            if (drCalCount["Exceptional"].ToString().Trim() != "7") //油車
                                            {
                                                if (drCalCount["Class"].ToString().Trim() == "甲") //大巴
                                                {
                                                    vTempFloatA = double.Parse(dtTarget.Rows[i]["HighWayCount_OA"].ToString().Trim());
                                                    dtTarget.Rows[i]["HighWayCount_OA"] = (vHasOldData) ? vTempFloatA : (vTempFloatA + vTempFloat);
                                                }
                                                else if (drCalCount["Class"].ToString().Trim() == "乙") //中巴
                                                {
                                                    vTempFloatB = double.Parse(dtTarget.Rows[i]["HighWayCount_OB"].ToString().Trim());
                                                    dtTarget.Rows[i]["HighWayCount_OB"] = (vHasOldData) ? vTempFloatB : (vTempFloatB + vTempFloat);
                                                }
                                                vTempFloatT = double.Parse(dtTarget.Rows[i]["HighWayCount_OT"].ToString().Trim());
                                                dtTarget.Rows[i]["HighWayCount_OT"] = (vHasOldData) ? vTempFloatT : (vTempFloatT + vTempFloat);
                                            }
                                            else //電車
                                            {
                                                if (drCalCount["Class"].ToString().Trim() == "甲") //大巴
                                                {
                                                    vTempFloatA = double.Parse(dtTarget.Rows[i]["HighWayCount_EA"].ToString().Trim());
                                                    dtTarget.Rows[i]["HighWayCount_EA"] = (vHasOldData) ? vTempFloatA : (vTempFloatA + vTempFloat);
                                                }
                                                else if (drCalCount["Class"].ToString().Trim() == "乙") //中巴
                                                {
                                                    vTempFloatB = double.Parse(dtTarget.Rows[i]["HighWayCount_EB"].ToString().Trim());
                                                    dtTarget.Rows[i]["HighWayCount_EB"] = (vHasOldData) ? vTempFloatB : (vTempFloatB + vTempFloat);
                                                }
                                                vTempFloatT = double.Parse(dtTarget.Rows[i]["HighWayCount_ET"].ToString().Trim());
                                                dtTarget.Rows[i]["HighWayCount_ET"] = (vHasOldData) ? vTempFloatT : (vTempFloatT + vTempFloat);
                                            }
                                        }
                                        else if (drCalCount["Point"].ToString().Trim().ToUpper() == "A2") //班車
                                        {
                                            if (drCalCount["Exceptional"].ToString().Trim() != "7") //油車
                                            {
                                                if (drCalCount["Class"].ToString().Trim() == "甲") //大巴
                                                {
                                                    vTempFloatA = double.Parse(dtTarget.Rows[i]["RuleCount_OA"].ToString().Trim());
                                                    dtTarget.Rows[i]["RuleCount_OA"] = (vHasOldData) ? vTempFloatA : (vTempFloatA + vTempFloat);
                                                }
                                                else if (drCalCount["Class"].ToString().Trim() == "乙") //中巴
                                                {
                                                    vTempFloatB = double.Parse(dtTarget.Rows[i]["RuleCount_OB"].ToString().Trim());
                                                    dtTarget.Rows[i]["RuleCount_OB"] = (vHasOldData) ? vTempFloatB : (vTempFloatB + vTempFloat);
                                                }
                                                vTempFloatT = double.Parse(dtTarget.Rows[i]["RuleCount_OT"].ToString().Trim());
                                                dtTarget.Rows[i]["RuleCount_OT"] = (vHasOldData) ? vTempFloatT : (vTempFloatT + vTempFloat);
                                            }
                                            else //電車
                                            {
                                                if (drCalCount["Class"].ToString().Trim() == "甲") //大巴
                                                {
                                                    vTempFloatA = double.Parse(dtTarget.Rows[i]["RuleCount_EA"].ToString().Trim());
                                                    dtTarget.Rows[i]["RuleCount_EA"] = (vHasOldData) ? vTempFloatA : (vTempFloatA + vTempFloat);
                                                }
                                                else if (drCalCount["Class"].ToString().Trim() == "乙") //中巴
                                                {
                                                    vTempFloatB = double.Parse(dtTarget.Rows[i]["RuleCount_EB"].ToString().Trim());
                                                    dtTarget.Rows[i]["RuleCount_EB"] = (vHasOldData) ? vTempFloatB : (vTempFloatB + vTempFloat);
                                                }
                                                vTempFloatT = double.Parse(dtTarget.Rows[i]["RuleCount_ET"].ToString().Trim());
                                                dtTarget.Rows[i]["RuleCount_ET"] = (vHasOldData) ? vTempFloatT : (vTempFloatT + vTempFloat);
                                            }
                                        }
                                        else if (drCalCount["Point"].ToString().Trim().ToUpper() == "B2") //公車
                                        {
                                            if (drCalCount["Exceptional"].ToString().Trim() != "7") //油車
                                            {
                                                if (drCalCount["Class"].ToString().Trim() == "甲") //大巴
                                                {
                                                    vTempFloatA = double.Parse(dtTarget.Rows[i]["BusCount_OA"].ToString().Trim());
                                                    dtTarget.Rows[i]["BusCount_OA"] = (vHasOldData) ? vTempFloatA : (vTempFloatA + vTempFloat);
                                                }
                                                else if (drCalCount["Class"].ToString().Trim() == "乙") //中巴
                                                {
                                                    vTempFloatB = double.Parse(dtTarget.Rows[i]["BusCount_OB"].ToString().Trim());
                                                    dtTarget.Rows[i]["BusCount_OB"] = (vHasOldData) ? vTempFloatB : (vTempFloatB + vTempFloat);
                                                }
                                                vTempFloatT = double.Parse(dtTarget.Rows[i]["BusCount_OT"].ToString().Trim());
                                                dtTarget.Rows[i]["BusCount_OT"] = (vHasOldData) ? vTempFloatT : (vTempFloatT + vTempFloat);
                                            }
                                            else //電車
                                            {
                                                if (drCalCount["Class"].ToString().Trim() == "甲") //大巴
                                                {
                                                    vTempFloatA = double.Parse(dtTarget.Rows[i]["BusCount_EA"].ToString().Trim());
                                                    dtTarget.Rows[i]["BusCount_EA"] = (vHasOldData) ? vTempFloatA : (vTempFloatA + vTempFloat);
                                                }
                                                else if (drCalCount["Class"].ToString().Trim() == "乙") //中巴
                                                {
                                                    vTempFloatB = double.Parse(dtTarget.Rows[i]["BusCount_EB"].ToString().Trim());
                                                    dtTarget.Rows[i]["BusCount_EB"] = (vHasOldData) ? vTempFloatB : (vTempFloatB + vTempFloat);
                                                }
                                                vTempFloatT = double.Parse(dtTarget.Rows[i]["BusCount_ET"].ToString().Trim());
                                                dtTarget.Rows[i]["BusCount_ET"] = (vHasOldData) ? vTempFloatT : (vTempFloatT + vTempFloat);
                                            }
                                        }
                                        else if (drCalCount["Point"].ToString().Trim().ToUpper() == "C1") //遊覽車
                                        {
                                            if (drCalCount["Exceptional"].ToString().Trim() != "7") //油車
                                            {
                                                if (drCalCount["Class"].ToString().Trim() == "甲") //大巴
                                                {
                                                    vTempFloatA = double.Parse(dtTarget.Rows[i]["TourCount_OA"].ToString().Trim());
                                                    dtTarget.Rows[i]["TourCount_OA"] = (vHasOldData) ? vTempFloatA : (vTempFloatA + vTempFloat);
                                                }
                                                else if (drCalCount["Class"].ToString().Trim() == "乙") //中巴
                                                {
                                                    vTempFloatB = double.Parse(dtTarget.Rows[i]["TourCount_OB"].ToString().Trim());
                                                    dtTarget.Rows[i]["TourCount_OB"] = (vHasOldData) ? vTempFloatB : (vTempFloatB + vTempFloat);
                                                }
                                                vTempFloatT = double.Parse(dtTarget.Rows[i]["TourCount_OT"].ToString().Trim());
                                                dtTarget.Rows[i]["TourCount_OT"] = (vHasOldData) ? vTempFloatT : (vTempFloatT + vTempFloat);
                                            }
                                            else //電車
                                            {
                                                if (drCalCount["Class"].ToString().Trim() == "甲") //大巴
                                                {
                                                    vTempFloatA = double.Parse(dtTarget.Rows[i]["TourCount_EA"].ToString().Trim());
                                                    dtTarget.Rows[i]["TourCount_EA"] = (vHasOldData) ? vTempFloatA : (vTempFloatA + vTempFloat);
                                                }
                                                else if (drCalCount["Class"].ToString().Trim() == "乙") //中巴
                                                {
                                                    vTempFloatB = double.Parse(dtTarget.Rows[i]["TourCount_EB"].ToString().Trim());
                                                    dtTarget.Rows[i]["TourCount_EB"] = (vHasOldData) ? vTempFloatB : (vTempFloatB + vTempFloat);
                                                }
                                                vTempFloatT = double.Parse(dtTarget.Rows[i]["TourCount_ET"].ToString().Trim());
                                                dtTarget.Rows[i]["TourCount_ET"] = (vHasOldData) ? vTempFloatT : (vTempFloatT + vTempFloat);
                                            }
                                        }
                                        else //其他
                                        {
                                            if (drCalCount["Exceptional"].ToString().Trim() != "7") //油車
                                            {
                                                if (drCalCount["Class"].ToString().Trim() == "甲") //大巴
                                                {
                                                    vTempFloatA = double.Parse(dtTarget.Rows[i]["OtherCount_OA"].ToString().Trim());
                                                    dtTarget.Rows[i]["OtherCount_OA"] = (vHasOldData) ? vTempFloatA : (vTempFloatA + vTempFloat);
                                                }
                                                else if (drCalCount["Class"].ToString().Trim() == "乙") //中巴
                                                {
                                                    vTempFloatB = double.Parse(dtTarget.Rows[i]["OtherCount_OB"].ToString().Trim());
                                                    dtTarget.Rows[i]["OtherCount_OB"] = (vHasOldData) ? vTempFloatB : (vTempFloatB + vTempFloat);
                                                }
                                                vTempFloatT = double.Parse(dtTarget.Rows[i]["OtherCount_OT"].ToString().Trim());
                                                dtTarget.Rows[i]["OtherCount_OT"] = (vHasOldData) ? vTempFloatT : (vTempFloatT + vTempFloat);
                                            }
                                            else //電車
                                            {
                                                if (drCalCount["Class"].ToString().Trim() == "甲") //大巴
                                                {
                                                    vTempFloatA = double.Parse(dtTarget.Rows[i]["OtherCount_EA"].ToString().Trim());
                                                    dtTarget.Rows[i]["OtherCount_EA"] = (vHasOldData) ? vTempFloatA : (vTempFloatA + vTempFloat);
                                                }
                                                else if (drCalCount["Class"].ToString().Trim() == "乙") //中巴
                                                {
                                                    vTempFloatB = double.Parse(dtTarget.Rows[i]["OtherCount_EB"].ToString().Trim());
                                                    dtTarget.Rows[i]["OtherCount_EB"] = (vHasOldData) ? vTempFloatB : (vTempFloatB + vTempFloat);
                                                }
                                                vTempFloatT = double.Parse(dtTarget.Rows[i]["OtherCount_ET"].ToString().Trim());
                                                dtTarget.Rows[i]["OtherCount_ET"] = (vHasOldData) ? vTempFloatT : (vTempFloatT + vTempFloat);
                                            }
                                        }
                                    }
                                }
                                //計算車輛數小計
                                dtTarget.Rows[i]["CarSum_OA"] = double.Parse(dtTarget.Rows[i]["RuleCount_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["HighWayCount_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusCount_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TourCount_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["OtherCount_OA"].ToString().Trim());
                                dtTarget.Rows[i]["CarSum_EA"] = double.Parse(dtTarget.Rows[i]["RuleCount_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["HighWayCount_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusCount_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TourCount_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["OtherCount_EA"].ToString().Trim());
                                dtTarget.Rows[i]["CarSum_OB"] = double.Parse(dtTarget.Rows[i]["RuleCount_OB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["HighWayCount_OB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusCount_OB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TourCount_OB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["OtherCount_OB"].ToString().Trim());
                                dtTarget.Rows[i]["CarSum_EB"] = double.Parse(dtTarget.Rows[i]["RuleCount_EB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["HighWayCount_EB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusCount_EB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TourCount_EB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["OtherCount_EB"].ToString().Trim());
                                dtTarget.Rows[i]["CarSum_OT"] = double.Parse(dtTarget.Rows[i]["RuleCount_OT"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["HighWayCount_OT"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusCount_OT"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TourCount_OT"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["OtherCount_OT"].ToString().Trim());
                                dtTarget.Rows[i]["CarSum_ET"] = double.Parse(dtTarget.Rows[i]["RuleCount_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["HighWayCount_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusCount_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TourCount_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["OtherCount_ET"].ToString().Trim());
                                dtTarget.Rows[i]["CarCount"] = double.Parse(dtTarget.Rows[i]["CarSum_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["CarSum_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["CarSum_OB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["CarSum_EB"].ToString().Trim());
                                //再計算里程數
                                vTempSelectStr = "select r.DepNo, Class, Exceptional, " + Environment.NewLine +
                                                 "       case when d.depGroup = '1' then(round(sum(CBus1Km), 1) + round(sum(CBus4Km), 1)) else round(sum(CBus1Km), 1) end as RuleKMS, " + Environment.NewLine +
                                                 "       case when d.depGroup = '2' then(round(sum(CBusCKm), 1) + round(sum(CBusNKm), 1) + round(sum(CBus4Km), 1)) else (round(sum(CBusCKm), 1) + round(sum(CBusNKm), 1)) end as BusKMS, " + Environment.NewLine +
                                                 "       0 as UnionKMS, round(sum(CBus3Km), 1) as SpecKMS, round(sum(CRentTraKm), 1) as TransKMS, round(sum(CRentAKm), 1) as RentKMS, " + Environment.NewLine +
                                                 "       round(sum(CRentBKm), 1) as TourKMS, round(sum(CBus5Km), 1) as NoneBusiKMS " + Environment.NewLine +
                                                 "  from RsCarMonth r left join Car_InfoA a on a.Car_ID = r.Car_ID " + Environment.NewLine +
                                                 "                    left join Department d on d.DepNo = r.DepNo " + Environment.NewLine +
                                                 " where r.budate >= '" + eCalYM_Temp.ToString("yyyy/MM/dd") + "' and budate <= '" + eCalYMLast_Temp.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                                                 "   and r.depno = '" + vDepNo + "' " + Environment.NewLine +
                                                 " group by r.DepNo, budate, Class, Exceptional, d.DepGroup";
                                using (SqlConnection connCalKMS = new SqlConnection(vConnStr))
                                {
                                    SqlCommand cmdCalKMS = new SqlCommand(vTempSelectStr, connCalKMS);
                                    connCalKMS.Open();
                                    SqlDataReader drCalKMS = cmdCalKMS.ExecuteReader();
                                    while (drCalKMS.Read())
                                    {
                                        if (drCalKMS["Exceptional"].ToString().Trim() != "7") //油車
                                        {
                                            if (drCalKMS["Class"].ToString().Trim() == "甲") //大巴
                                            {
                                                //班車里程
                                                vTempRuleKMS = double.Parse(drCalKMS["RuleKMS"].ToString().Trim());
                                                vTempRuleKMS_T = double.Parse(dtTarget.Rows[i]["RuleKMS_OA"].ToString().Trim());
                                                dtTarget.Rows[i]["RuleKMS_OA"] = (vTempRuleKMS_T + vTempRuleKMS);
                                                //公車里程
                                                vTempBusKMS = double.Parse(drCalKMS["BusKMS"].ToString().Trim());
                                                vTempBusKMS_T = double.Parse(dtTarget.Rows[i]["BusKMS_OA"].ToString().Trim());
                                                dtTarget.Rows[i]["BusKMS_OA"] = (vTempBusKMS_T + vTempBusKMS);
                                                //區間租車里程
                                                vTempRentKMS = double.Parse(drCalKMS["RentKMS"].ToString().Trim());
                                                vTempRentKMS_T = double.Parse(dtTarget.Rows[i]["RentKMS_OA"].ToString().Trim());
                                                dtTarget.Rows[i]["RentKMS_OA"] = (vTempRentKMS + vTempRentKMS_T);
                                                //交通車里程
                                                vTempTransKMS = double.Parse(drCalKMS["TransKMS"].ToString().Trim());
                                                vTempTransKMS_T = double.Parse(dtTarget.Rows[i]["TransKMS_OA"].ToString().Trim());
                                                dtTarget.Rows[i]["TransKMS_OA"] = (vTempTransKMS_T + vTempTransKMS);
                                                //遊覽車里程
                                                vTempTourKMS = double.Parse(drCalKMS["TourKMS"].ToString().Trim());
                                                vTempTourKMS_T = double.Parse(dtTarget.Rows[i]["TourKMS_OA"].ToString().Trim());
                                                dtTarget.Rows[i]["TourKMS_OA"] = (vTempTourKMS_T + vTempTourKMS);
                                                //非營運里程
                                                vTempNoneBusiKMS = double.Parse(drCalKMS["NoneBusiKMS"].ToString().Trim());
                                                vTempNoneBusiKMS_T = double.Parse(dtTarget.Rows[i]["NoneBusiKMS_OA"].ToString().Trim());
                                                dtTarget.Rows[i]["NoneBusiKMS_OA"] = (vTempNoneBusiKMS_T + vTempNoneBusiKMS);
                                                //專車里程
                                                vTempSpecKMS = double.Parse(drCalKMS["SpecKMS"].ToString().Trim());
                                                vTempSpecKMS_T = double.Parse(dtTarget.Rows[i]["SpecKMS_OA"].ToString().Trim());
                                                dtTarget.Rows[i]["SpecKMS_OA"] = (vTempSpecKMS_T + vTempSpecKMS);
                                            }
                                            else if (drCalKMS["Class"].ToString().Trim() == "乙") //中巴
                                            {
                                                //班車里程
                                                vTempRuleKMS = double.Parse(drCalKMS["RuleKMS"].ToString().Trim());
                                                vTempRuleKMS_T = double.Parse(dtTarget.Rows[i]["RuleKMS_OB"].ToString().Trim());
                                                dtTarget.Rows[i]["RuleKMS_OB"] = (vTempRuleKMS_T + vTempRuleKMS);
                                                //公車里程
                                                vTempBusKMS = double.Parse(drCalKMS["BusKMS"].ToString().Trim());
                                                vTempBusKMS_T = double.Parse(dtTarget.Rows[i]["BusKMS_OB"].ToString().Trim());
                                                dtTarget.Rows[i]["BusKMS_OB"] = (vTempBusKMS_T + vTempBusKMS);
                                                //區間租車里程
                                                vTempRentKMS = double.Parse(drCalKMS["RentKMS"].ToString().Trim());
                                                vTempRentKMS_T = double.Parse(dtTarget.Rows[i]["RentKMS_OB"].ToString().Trim());
                                                dtTarget.Rows[i]["RentKMS_OB"] = (vTempRentKMS + vTempRentKMS_T);
                                                //交通車里程
                                                vTempTransKMS = double.Parse(drCalKMS["TransKMS"].ToString().Trim());
                                                vTempTransKMS_T = double.Parse(dtTarget.Rows[i]["TransKMS_OB"].ToString().Trim());
                                                dtTarget.Rows[i]["TransKMS_OB"] = (vTempTransKMS_T + vTempTransKMS);
                                                //遊覽車里程
                                                vTempTourKMS = double.Parse(drCalKMS["TourKMS"].ToString().Trim());
                                                vTempTourKMS_T = double.Parse(dtTarget.Rows[i]["TourKMS_OB"].ToString().Trim());
                                                dtTarget.Rows[i]["TourKMS_OB"] = (vTempTourKMS_T + vTempTourKMS);
                                                //非營運里程
                                                vTempNoneBusiKMS = double.Parse(drCalKMS["NoneBusiKMS"].ToString().Trim());
                                                vTempNoneBusiKMS_T = double.Parse(dtTarget.Rows[i]["NoneBusiKMS_OB"].ToString().Trim());
                                                dtTarget.Rows[i]["NoneBusiKMS_OB"] = (vTempNoneBusiKMS_T + vTempNoneBusiKMS);
                                                //專車里程
                                                vTempSpecKMS = double.Parse(drCalKMS["SpecKMS"].ToString().Trim());
                                                vTempSpecKMS_T = double.Parse(dtTarget.Rows[i]["SpecKMS_OB"].ToString().Trim());
                                                dtTarget.Rows[i]["SpecKMS_OB"] = (vTempSpecKMS_T + vTempSpecKMS);
                                            }
                                        }
                                        else //電車
                                        {
                                            if (drCalKMS["Class"].ToString().Trim() == "甲") //大巴
                                            {
                                                //班車里程
                                                vTempRuleKMS = double.Parse(drCalKMS["RuleKMS"].ToString().Trim());
                                                vTempRuleKMS_T = double.Parse(dtTarget.Rows[i]["RuleKMS_EA"].ToString().Trim());
                                                dtTarget.Rows[i]["RuleKMS_EA"] = (vTempRuleKMS_T + vTempRuleKMS);
                                                //公車里程
                                                vTempBusKMS = double.Parse(drCalKMS["BusKMS"].ToString().Trim());
                                                vTempBusKMS_T = double.Parse(dtTarget.Rows[i]["BusKMS_EA"].ToString().Trim());
                                                dtTarget.Rows[i]["BusKMS_EA"] = (vTempBusKMS_T + vTempBusKMS);
                                                //區間租車里程
                                                vTempRentKMS = double.Parse(drCalKMS["RentKMS"].ToString().Trim());
                                                vTempRentKMS_T = double.Parse(dtTarget.Rows[i]["RentKMS_EA"].ToString().Trim());
                                                dtTarget.Rows[i]["RentKMS_EA"] = (vTempRentKMS + vTempRentKMS_T);
                                                //交通車里程
                                                vTempTransKMS = double.Parse(drCalKMS["TransKMS"].ToString().Trim());
                                                vTempTransKMS_T = double.Parse(dtTarget.Rows[i]["TransKMS_EA"].ToString().Trim());
                                                dtTarget.Rows[i]["TransKMS_EA"] = (vTempTransKMS_T + vTempTransKMS);
                                                //遊覽車里程
                                                vTempTourKMS = double.Parse(drCalKMS["TourKMS"].ToString().Trim());
                                                vTempTourKMS_T = double.Parse(dtTarget.Rows[i]["TourKMS_EA"].ToString().Trim());
                                                dtTarget.Rows[i]["TourKMS_EA"] = (vTempTourKMS_T + vTempTourKMS);
                                                //非營運里程
                                                vTempNoneBusiKMS = double.Parse(drCalKMS["NoneBusiKMS"].ToString().Trim());
                                                vTempNoneBusiKMS_T = double.Parse(dtTarget.Rows[i]["NoneBusiKMS_EA"].ToString().Trim());
                                                dtTarget.Rows[i]["NoneBusiKMS_EA"] = (vTempNoneBusiKMS_T + vTempNoneBusiKMS);
                                                //專車里程
                                                vTempSpecKMS = double.Parse(drCalKMS["SpecKMS"].ToString().Trim());
                                                vTempSpecKMS_T = double.Parse(dtTarget.Rows[i]["SpecKMS_EA"].ToString().Trim());
                                                dtTarget.Rows[i]["SpecKMS_EA"] = (vTempSpecKMS_T + vTempSpecKMS);
                                            }
                                            else if (drCalKMS["Class"].ToString().Trim() == "乙") //中巴
                                            {
                                                //班車里程
                                                vTempRuleKMS = double.Parse(drCalKMS["RuleKMS"].ToString().Trim());
                                                vTempRuleKMS_T = double.Parse(dtTarget.Rows[i]["RuleKMS_EB"].ToString().Trim());
                                                dtTarget.Rows[i]["RuleKMS_EB"] = (vTempRuleKMS_T + vTempRuleKMS);
                                                //公車里程
                                                vTempBusKMS = double.Parse(drCalKMS["BusKMS"].ToString().Trim());
                                                vTempBusKMS_T = double.Parse(dtTarget.Rows[i]["BusKMS_EB"].ToString().Trim());
                                                dtTarget.Rows[i]["BusKMS_EB"] = (vTempBusKMS_T + vTempBusKMS);
                                                //區間租車里程
                                                vTempRentKMS = double.Parse(drCalKMS["RentKMS"].ToString().Trim());
                                                vTempRentKMS_T = double.Parse(dtTarget.Rows[i]["RentKMS_EB"].ToString().Trim());
                                                dtTarget.Rows[i]["RentKMS_EB"] = (vTempRentKMS + vTempRentKMS_T);
                                                //交通車里程
                                                vTempTransKMS = double.Parse(drCalKMS["TransKMS"].ToString().Trim());
                                                vTempTransKMS_T = double.Parse(dtTarget.Rows[i]["TransKMS_EB"].ToString().Trim());
                                                dtTarget.Rows[i]["TransKMS_EB"] = (vTempTransKMS_T + vTempTransKMS);
                                                //遊覽車里程
                                                vTempTourKMS = double.Parse(drCalKMS["TourKMS"].ToString().Trim());
                                                vTempTourKMS_T = double.Parse(dtTarget.Rows[i]["TourKMS_EB"].ToString().Trim());
                                                dtTarget.Rows[i]["TourKMS_EB"] = (vTempTourKMS_T + vTempTourKMS);
                                                //非營運里程
                                                vTempNoneBusiKMS = double.Parse(drCalKMS["NoneBusiKMS"].ToString().Trim());
                                                vTempNoneBusiKMS_T = double.Parse(dtTarget.Rows[i]["NoneBusiKMS_EB"].ToString().Trim());
                                                dtTarget.Rows[i]["NoneBusiKMS_EB"] = (vTempNoneBusiKMS_T + vTempNoneBusiKMS);
                                                //專車里程
                                                vTempSpecKMS = double.Parse(drCalKMS["SpecKMS"].ToString().Trim());
                                                vTempSpecKMS_T = double.Parse(dtTarget.Rows[i]["SpecKMS_EB"].ToString().Trim());
                                                dtTarget.Rows[i]["SpecKMS_EB"] = (vTempSpecKMS_T + vTempSpecKMS);
                                            }
                                        }
                                    }
                                }
                                //排除觀光路線的專車里程，歸屬回公車里程
                                vTempSelectStr = "select sum(b.ActualKM) TotalKM, c.Class, c.Exceptional " + Environment.NewLine +
                                                 "  from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                                 "                   left join CAR_infoA c on c.Car_ID = b.Car_ID " + Environment.NewLine +
                                                 "                   left join Lines d on d.LinesNo = b.LinesNo " + Environment.NewLine +
                                                 " where a.Budate between '" + eCalYM_Temp.ToString("yyyy/MM/dd") + "' and '" + eCalYMLast_Temp.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                                                 "   and a.DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                                 "   and b.CarType = '2' and isnull(b.ReduceReason,'') = '' " + Environment.NewLine +
                                                 "   and d.BSDBSPKM = 'V' " + Environment.NewLine +
                                                 " Group by c.Class, c.Exceptional";
                                using (SqlConnection connModifyKMS = new SqlConnection(vConnStr))
                                {
                                    SqlCommand cmdModifyKMS = new SqlCommand(vTempSelectStr, connModifyKMS);
                                    connModifyKMS.Open();
                                    SqlDataReader drModifyKMS = cmdModifyKMS.ExecuteReader();
                                    while (drModifyKMS.Read())
                                    {
                                        vTempFloat = double.Parse(drModifyKMS["TotalKM"].ToString().Trim());
                                        if (drModifyKMS["Exceptional"].ToString().Trim() != "7") //油車
                                        {
                                            if (drModifyKMS["Class"].ToString().Trim() == "甲") //大巴
                                            {
                                                vTempSpecKMS = double.Parse(dtTarget.Rows[i]["SpecKMS_OA"].ToString().Trim());
                                                vTempBusKMS = double.Parse(dtTarget.Rows[i]["BusKMS_OA"].ToString().Trim());
                                                dtTarget.Rows[i]["SpecKMS_OA"] = vTempSpecKMS - vTempFloat;
                                                dtTarget.Rows[i]["BusKMS_OA"] = vTempBusKMS + vTempFloat;
                                            }
                                            else if (drModifyKMS["Class"].ToString().Trim() == "乙") //中巴
                                            {
                                                vTempSpecKMS = double.Parse(dtTarget.Rows[i]["SpecKMS_OB"].ToString().Trim());
                                                vTempBusKMS = double.Parse(dtTarget.Rows[i]["BusKMS_OB"].ToString().Trim());
                                                dtTarget.Rows[i]["SpecKMS_OB"] = vTempSpecKMS - vTempFloat;
                                                dtTarget.Rows[i]["BusKMS_OB"] = vTempBusKMS + vTempFloat;
                                            }
                                        }
                                        else //電車
                                        {
                                            if (drModifyKMS["Class"].ToString().Trim() == "甲") //大巴
                                            {
                                                vTempSpecKMS = double.Parse(dtTarget.Rows[i]["SpecKMS_EA"].ToString().Trim());
                                                vTempBusKMS = double.Parse(dtTarget.Rows[i]["BusKMS_EA"].ToString().Trim());
                                                dtTarget.Rows[i]["SpecKMS_EA"] = vTempSpecKMS - vTempFloat;
                                                dtTarget.Rows[i]["BusKMS_EA"] = vTempBusKMS + vTempFloat;
                                            }
                                            else if (drModifyKMS["Class"].ToString().Trim() == "乙") //中巴
                                            {
                                                vTempSpecKMS = double.Parse(dtTarget.Rows[i]["SpecKMS_EB"].ToString().Trim());
                                                vTempBusKMS = double.Parse(dtTarget.Rows[i]["BusKMS_EB"].ToString().Trim());
                                                dtTarget.Rows[i]["SpecKMS_EB"] = vTempSpecKMS - vTempFloat;
                                                dtTarget.Rows[i]["BusKMS_EB"] = vTempBusKMS + vTempFloat;
                                            }
                                        }
                                    }
                                }
                                //計算各種里程總數
                                //油料大型巴士總里程
                                dtTarget.Rows[i]["TotalKMS_OA"] = double.Parse(dtTarget.Rows[i]["RuleKMS_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusKMS_OA"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["SpecKMS_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["RentKMS_OA"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["TransKMS_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["NoneBusiKMS_OA"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["TourKMS_OA"].ToString().Trim());
                                //電能大型巴士總里程
                                dtTarget.Rows[i]["TotalKMS_EA"] = double.Parse(dtTarget.Rows[i]["RuleKMS_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusKMS_EA"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["SpecKMS_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["RentKMS_EA"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["TransKMS_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["NoneBusiKMS_EA"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["TourKMS_EA"].ToString().Trim());
                                //油料中型巴士總里程
                                dtTarget.Rows[i]["TotalKMS_OB"] = double.Parse(dtTarget.Rows[i]["RuleKMS_OB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusKMS_OB"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["SpecKMS_OB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["RentKMS_OB"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["TransKMS_OB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["NoneBusiKMS_OB"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["TourKMS_OB"].ToString().Trim());
                                //電能中型巴士總里程
                                dtTarget.Rows[i]["TotalKMS_EB"] = double.Parse(dtTarget.Rows[i]["RuleKMS_EB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusKMS_EB"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["SpecKMS_EB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["RentKMS_EB"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["TransKMS_EB"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["NoneBusiKMS_EB"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["TourKMS_EB"].ToString().Trim());
                                //油料車輛總里程
                                dtTarget.Rows[i]["TotalKMS_OT"] = double.Parse(dtTarget.Rows[i]["RuleKMS_OT"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusKMS_OT"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["SpecKMS_OT"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["RentKMS_OT"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["TransKMS_OT"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["NoneBusiKMS_OT"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["TourKMS_OT"].ToString().Trim());
                                //電能車輛總里程
                                dtTarget.Rows[i]["TotalKMS_ET"] = double.Parse(dtTarget.Rows[i]["RuleKMS_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusKMS_ET"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["SpecKMS_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["RentKMS_ET"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["TransKMS_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["NoneBusiKMS_ET"].ToString().Trim()) +
                                                                  double.Parse(dtTarget.Rows[i]["TourKMS_ET"].ToString().Trim());
                                //油班車里程
                                dtTarget.Rows[i]["RuleKMS_OT"] = double.Parse(dtTarget.Rows[i]["RuleKMS_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["RuleKMS_OB"].ToString().Trim());
                                //電班車里程
                                dtTarget.Rows[i]["RuleKMS_ET"] = double.Parse(dtTarget.Rows[i]["RuleKMS_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["RuleKMS_EB"].ToString().Trim());
                                //油公車里程
                                dtTarget.Rows[i]["BusKMS_OT"] = double.Parse(dtTarget.Rows[i]["BusKMS_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusKMS_OB"].ToString().Trim());
                                //電公車里程
                                dtTarget.Rows[i]["BusKMS_ET"] = double.Parse(dtTarget.Rows[i]["BusKMS_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusKMS_EB"].ToString().Trim());
                                //油專車里程
                                dtTarget.Rows[i]["SpecKMS_OT"] = double.Parse(dtTarget.Rows[i]["SpecKMS_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["SpecKMS_OB"].ToString().Trim());
                                //電專車里程
                                dtTarget.Rows[i]["SpecKMS_ET"] = double.Parse(dtTarget.Rows[i]["SpecKMS_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["SpecKMS_EB"].ToString().Trim());
                                //油租車里程
                                dtTarget.Rows[i]["RentKMS_OT"] = double.Parse(dtTarget.Rows[i]["RentKMS_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["RentKMS_OB"].ToString().Trim());
                                //電租車里程
                                dtTarget.Rows[i]["RentKMS_ET"] = double.Parse(dtTarget.Rows[i]["RentKMS_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["RentKMS_EB"].ToString().Trim());
                                //油文通車里程
                                dtTarget.Rows[i]["TransKMS_OT"] = double.Parse(dtTarget.Rows[i]["TransKMS_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TransKMS_OB"].ToString().Trim());
                                //電交通車里程
                                dtTarget.Rows[i]["TransKMS_ET"] = double.Parse(dtTarget.Rows[i]["TransKMS_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TransKMS_EB"].ToString().Trim());
                                //油遊覽車里程
                                dtTarget.Rows[i]["TourKMS_OT"] = double.Parse(dtTarget.Rows[i]["TourKMS_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TourKMS_OB"].ToString().Trim());
                                //電遊覽車里程
                                dtTarget.Rows[i]["TourKMS_ET"] = double.Parse(dtTarget.Rows[i]["TourKMS_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TourKMS_EB"].ToString().Trim());
                                //油非營運里程
                                dtTarget.Rows[i]["NoneBusiKMS_OT"] = double.Parse(dtTarget.Rows[i]["NoneBusiKMS_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["NoneBusiKMS_OB"].ToString().Trim());
                                //電非營運里程
                                dtTarget.Rows[i]["NoneBusiKMS_ET"] = double.Parse(dtTarget.Rows[i]["NoneBusiKMS_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["NoneBusiKMS_EB"].ToString().Trim());
                                //油車里程
                                dtTarget.Rows[i]["TotalKMS_OT"] = double.Parse(dtTarget.Rows[i]["TotalKMS_OA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TotalKMS_OB"].ToString().Trim());
                                //電車里程
                                dtTarget.Rows[i]["TotalKMS_ET"] = double.Parse(dtTarget.Rows[i]["TotalKMS_EA"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TotalKMS_EB"].ToString().Trim());
                                //班車
                                dtTarget.Rows[i]["KMS_Rule"] = double.Parse(dtTarget.Rows[i]["RuleKMS_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["RuleKMS_OT"].ToString().Trim());
                                //公車
                                dtTarget.Rows[i]["KMS_BUS"] = double.Parse(dtTarget.Rows[i]["BusKMS_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["BusKMS_OT"].ToString().Trim());
                                //專車
                                dtTarget.Rows[i]["KMS_Spec"] = double.Parse(dtTarget.Rows[i]["SpecKMS_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["SpecKMS_OT"].ToString().Trim());
                                //交通車
                                dtTarget.Rows[i]["KMS_Trans"] = double.Parse(dtTarget.Rows[i]["TransKMS_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TransKMS_OT"].ToString().Trim());
                                //租車
                                dtTarget.Rows[i]["KMS_Rent"] = double.Parse(dtTarget.Rows[i]["RentKMS_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["RentKMS_OT"].ToString().Trim());
                                //遊覽車
                                dtTarget.Rows[i]["KMS_Tour"] = double.Parse(dtTarget.Rows[i]["TourKMS_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TourKMS_OT"].ToString().Trim());
                                //非營運里程
                                dtTarget.Rows[i]["KMS_NoneBusi"] = double.Parse(dtTarget.Rows[i]["NoneBusiKMS_ET"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["NoneBusiKMS_OT"].ToString().Trim());
                                //站總里程
                                dtTarget.Rows[i]["DriveRange"] = double.Parse(dtTarget.Rows[i]["TotalKMS_OT"].ToString().Trim()) + double.Parse(dtTarget.Rows[i]["TotalKMS_ET"].ToString().Trim());
                                //計算使用車輛數
                                vTempSelectStr = "select count(Car_ID) RCount " + Environment.NewLine +
                                                 "  from RsCarTotal " + Environment.NewLine +
                                                 " where DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                                 "   and BuDate between '" + eCalYM_Temp.ToString("yyyy/MM/dd") + " 00:00:00' and '" + eCalYMLast_Temp.ToString("yyyy/MM/dd") + " 23:59:59' ";
                                dtTarget.Rows[i]["UsedCars"] = double.Parse(PF.GetValue(vConnStr, vTempSelectStr, "RCount"));
                                //計算電子票證載客人數
                                vTempSelectStr = "select sum(HCount) as HCount " + Environment.NewLine +
                                                 "  from CPSOUTG " + Environment.NewLine +
                                                 " where DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                                 "   and [Date] between '" + eCalYM_Temp.ToString("yyyy/MM/dd") + " 00:00:00' and '" + eCalYMLast_Temp.ToString("yyyy/MM/dd") + " 23:59:59' ";

                                dtTarget.Rows[i]["ElecCardCount"] = double.TryParse(PF.GetValue(vConnStr, vTempSelectStr, "HCount"), out vTempFloat) ? vTempFloat : 0;
                                //計算各站配置員工數
                                vTempSelectStr = "select count(bb.EmpNo) as EmpCount,bb.Type,bb.DepNo " + Environment.NewLine +
                                                 "  from( " + Environment.NewLine +
                                                 "        select aa.EmpNo, aa.Type,case when isnull(aa.DepNo_S, '') <> '' then aa.DepNo_S else aa.DepNo end as DepNo " + Environment.NewLine +
                                                 "          from( " + Environment.NewLine +
                                                 "                select a.EmpNo, a.Name, a.Type, a.DepNo, " + Environment.NewLine +
                                                 "                       (select top 1 DepNo from Alters b where b.EmpNo = a.EmpNo and b.AlterType in ('B1', 'C1', 'D1', 'D11', 'E1') " + Environment.NewLine +
                                                 "                           and b.Inuredate <= '" + eCalYMLast_Temp.ToString("yyyy/MM/dd") + " 23:59:59' and isnull(IsAbort, 'N') = 'N' order by Inuredate DESC) as DepNo_S " + Environment.NewLine +
                                                 "                  from Employee a where a.EmpNo <> 'supervisor' and a.Assumeday <= '" + eCalYMLast_Temp.ToString("yyyy/MM/dd") + " 23:59:59' and(a.Leaveday is null or a.Leaveday > '" + eCalYMLast_Temp.ToString("yyyy/MM/dd") + " 00:00:00') " + Environment.NewLine +
                                                 "               ) aa " + Environment.NewLine +
                                                 "       ) bb where bb.DepNo = '" + vDepNo + "' group by bb.Type,bb.DepNo ";
                                using (SqlConnection connEmpCount = new SqlConnection(vConnStr))
                                {
                                    vEmpCount_M = 0.0;
                                    vEmpCount_F = 0.0;
                                    vEmpCount_D = 0.0;
                                    SqlCommand cmdEmpCount = new SqlCommand(vTempSelectStr, connEmpCount);
                                    connEmpCount.Open();
                                    SqlDataReader drEmpCount = cmdEmpCount.ExecuteReader();
                                    while (drEmpCount.Read())
                                    {
                                        if (drEmpCount["Type"].ToString().Trim() == "00")
                                        {
                                            vEmpCount_M = double.Parse(drEmpCount["EmpCount"].ToString().Trim());
                                        }
                                        else if (drEmpCount["Type"].ToString().Trim() == "10")
                                        {
                                            vEmpCount_F = double.Parse(drEmpCount["EmpCount"].ToString().Trim());
                                        }
                                        else if (drEmpCount["Type"].ToString().Trim() == "20")
                                        {
                                            vEmpCount_D = double.Parse(drEmpCount["EmpCount"].ToString().Trim());
                                        }
                                    }
                                    vEmpCount = vEmpCount_D + vEmpCount_F + vEmpCount_M;
                                    dtTarget.Rows[i]["EmpCount"] = vEmpCount;
                                    dtTarget.Rows[i]["EmpCount_M"] = vEmpCount_M;
                                    dtTarget.Rows[i]["EmpCount_F"] = vEmpCount_F;
                                    dtTarget.Rows[i]["EmpCount_D"] = vEmpCount_D;
                                }
                                //計算駕駛員加班時數
                                vODYM = eCalYM_Temp.Year.ToString("D4") + eCalYM_Temp.Month.ToString("D2");
                                vTempSelectStr = "select sum(TotalMIns - BasicMin) as OverMin from DriverOverDuty where OverYM = '" + vODYM + "' and DepNo = '" + vDepNo + "' ";
                                if (double.TryParse(PF.GetValue(vConnStr, vTempSelectStr, "OverMin"), out vTempFloat))
                                {
                                    vTempINT = (int)vTempFloat / 60;
                                    vDriverOD = (vTempFloat - (vTempINT * 60) >= 30) ? vTempINT + 1 : vTempINT;
                                }
                                else
                                {
                                    vDriverOD = 0;
                                }
                                dtTarget.Rows[i]["DriverOD"] = vDriverOD;
                                //計算拋錨次數
                                vTempSelectStr = "select count(b.AssignNoItem) RCount " + Environment.NewLine +
                                                 "  from RunSheetB b left join RunSheetA a on a.AssignNo = b.AssignNo " + Environment.NewLine +
                                                 " where a.DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                                 "   and a.BuDate between '" + eCalYM_Temp.ToString("yyyy/MM/dd") + " 00:00:00' and '" + eCalYMLast_Temp.ToString("yyyy/MM/dd") + " 23:59:59' " + Environment.NewLine +
                                                 "   and isnull(b.ReduceReason, '') = '7' ";
                                dtTarget.Rows[i]["BreakDown"] = double.TryParse(PF.GetValue(vConnStr, vTempSelectStr, "RCount"), out vTempFloat) ? vTempFloat : 0;
                            }
                            if ((dtTarget != null) && (dtTarget.Rows.Count > 0))
                            {
                                PF.ExecSQL(vConnStr, "delete CarCount where CalYM = '" + eCalYM_Temp.ToString("yyyy/MM/dd") + "' "); //先把原本的資料清除
                                using (SqlBulkCopy sbkTemp = new SqlBulkCopy(vConnStr, SqlBulkCopyOptions.UseInternalTransaction))
                                {
                                    try
                                    {
                                        //sbkTemp.DestinationTableName = "TLD_Data";
                                        sbkTemp.DestinationTableName = "CarCount";
                                        sbkTemp.WriteToServer(dtTarget);
                                        dtTarget.Clear();
                                        Response.Write("<Script language='Javascript'>");
                                        Response.Write("alert('資料計算完畢!')");
                                        Response.Write("</" + "Script>");
                                    }
                                    catch (Exception eMessage)
                                    {
                                        Response.Write("<Script language='Javascript'>");
                                        Response.Write("alert('" + eMessage.Message + "')");
                                        Response.Write("</" + "Script>");
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('計算年月格式錯誤！')");
                    Response.Write("</" + "Script>");
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先指定計算年月！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        private void OpenData()
        {
            string vSelectStr = "";
            if (eCalYM_Search.Text.Trim() != "")
            {
                //用 System.Linq 提供的函式 string.All() 來檢查字串是不是都是數字
                if ((eCalYM_Search.Text.Trim().Length == 5) && (eCalYM_Search.Text.Trim().All(char.IsDigit)))
                {
                    /*
                    vSelectStr = "SELECT DepNoYM, CalYM, c.DepNo, d.[Name] DepName, ModifyYM, ModifyMan, e.[Name] EmpName, CarCount, DriveRange, c.Remark, " + Environment.NewLine +
                                 "       KMS_BUS, KMS_Rent, KMS_Tour, KMS_Trans, KMS_Spec, RuleCount_OA, RuleCount_EA, RuleCount_OB, RuleCount_EB, RuleCount_OT, RuleCount_ET, " + Environment.NewLine +
                                 "       BusCount_OA, BusCount_EA, BusCount_OB, BusCount_EB, BusCount_OT, BusCount_ET, TourCount_OA, TourCount_EA, TourCount_OB, TourCount_EB, TourCount_OT, TourCount_ET, " + Environment.NewLine +
                                 "       CarSum_OA, CarSum_EA, CarSum_OB, CarSum_EB, CarSum_OT, CarSum_ET, RuleKMS_OA, RuleKMS_EA, RuleKMS_OB, RuleKMS_EB, RuleKMS_OT, RuleKMS_ET, KMS_Rule, " + Environment.NewLine +
                                 "       BusKMS_OA, BusKMS_EA, BusKMS_OB, BusKMS_EB, BusKMS_OT, BusKMS_ET, SpecKMS_OA, SpecKMS_EA, SpecKMS_OB, SpecKMS_EB, SpecKMS_OT, SpecKMS_ET, " + Environment.NewLine +
                                 "       RentKMS_OA, RentKMS_EA, RentKMS_OB, RentKMS_EB, RentKMS_OT, RentKMS_ET, TransKMS_OA, TransKMS_EA, TransKMS_OB, TransKMS_EB, TransKMS_OT, TransKMS_ET, " + Environment.NewLine +
                                 "       TourKMS_OA, TourKMS_EA, TourKMS_OB, TourKMS_EB, TourKMS_OT, TourKMS_ET, NoneBusiKMS_OA, NoneBusiKMS_EA, NoneBusiKMS_OB, NoneBusiKMS_EB, NoneBusiKMS_OT,NoneBusiKMS_ET, KMS_NoneBusi, " + Environment.NewLine +
                                 "       TotalKMS_OA, TotalKMS_EA, TotalKMS_OB, TotalKMS_EB, TotalKMS_OT, TotalKMS_ET, KMS_Total, HighWayCount_OA, HighWayCount_EA, HighWayCount_OB, HighWayCount_EB, HighWayCount_OT, HighWayCount_ET, " + Environment.NewLine +
                                 "       OtherCount_OA, OtherCount_EA, OtherCount_OB, OtherCount_EB, OtherCount_OT, OtherCount_ET, ModifyKMS, SpecIncome, UsedCars, ElecCardCount, EmpCount, DriverOD, EmpCount_M, EmpCount_F, EmpCount_D, BreakDown, " + Environment.NewLine +
                                 "       (CarSum_OA + CarSum_EA) CarSumA, (CarSum_OB + CarSum_EB) CarSumB " + Environment.NewLine +
                                 "  FROM dbo.CarCount c left join Department d on d.DepNo = c.DepNo " + Environment.NewLine +
                                 "                      left join Employee e on e.EmpNo = c.ModifyMan " + Environment.NewLine +
                                 " WHERE isnull(DepNoYM, '') like '" + eCalYM_Search.Text.Trim() + "%'" + Environment.NewLine +
                                 " ORDER BY c.DepNoYM ";
                    sdCarCount.SelectCommand = vSelectStr;
                    sdCarCount.DataBind();
                    fvCarCount.DataBind(); //*/
                    vSelectStr = "select c.CalYM, c.DepNoYM, c.DepNo, d.[Name] DepName, UsedCars, CarCount, EmpCount " + Environment.NewLine +
                                 "  from CarCount c left join Department d on d.DepNo = c.DepNo " + Environment.NewLine +
                                 " where isnull(DepNoYM, '') like '" + eCalYM_Search.Text.Trim() + "%'" + Environment.NewLine +
                                 " ORDER BY c.DepNoYM ";
                    dsCarCount_List.SelectCommand = vSelectStr;
                    dsCarCount_List.DataBind();
                    gvCarCount_List.DataBind();
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('計算年月格式錯誤！')");
                    Response.Write("</" + "Script>");
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先指定計算年月！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void eRuleCount_OA_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eRuleCount_OA = (TextBox)fvCarCount.FindControl("eRuleCount_OA_Edit");
            double vRuleCount_OA = 0.0;
            double.TryParse(eRuleCount_OA.Text.Trim(), out vRuleCount_OA);
            TextBox eHighWayCount_OA = (TextBox)fvCarCount.FindControl("eHighWayCount_OA_Edit");
            double vHighWayCount_OA = 0.0;
            double.TryParse(eHighWayCount_OA.Text.Trim(), out vHighWayCount_OA);
            TextBox eBusCount_OA = (TextBox)fvCarCount.FindControl("eBusCount_OA_Edit");
            double vBusCount_OA = 0.0;
            double.TryParse(eBusCount_OA.Text.Trim(), out vBusCount_OA);
            TextBox eTourCount_OA = (TextBox)fvCarCount.FindControl("eTourCount_OA_Edit");
            double vTourCount_OA = 0.0;
            double.TryParse(eTourCount_OA.Text.Trim(), out vTourCount_OA);
            TextBox eOtherCount_OA = (TextBox)fvCarCount.FindControl("eOtherCount_OA_Edit");
            double vOtherCount_OA = 0.0;
            double.TryParse(eOtherCount_OA.Text.Trim(), out vOtherCount_OA);
            TextBox eRuleCount_OB = (TextBox)fvCarCount.FindControl("eRuleCount_OB_Edit");
            double vRuleCount_OB = 0.0;
            double.TryParse(eRuleCount_OB.Text.Trim(), out vRuleCount_OB);
            TextBox eHighWayCount_OB = (TextBox)fvCarCount.FindControl("eHighWayCount_OB_Edit");
            double vHighWayCount_OB = 0.0;
            double.TryParse(eHighWayCount_OB.Text.Trim(), out vHighWayCount_OB);
            TextBox eBusCount_OB = (TextBox)fvCarCount.FindControl("eBusCount_OB_Edit");
            double vBusCount_OB = 0.0;
            double.TryParse(eBusCount_OB.Text.Trim(), out vBusCount_OB);
            TextBox eTourCount_OB = (TextBox)fvCarCount.FindControl("eTourCount_OB_Edit");
            double vTourCount_OB = 0.0;
            double.TryParse(eTourCount_OB.Text.Trim(), out vTourCount_OB);
            TextBox eOtherCount_OB = (TextBox)fvCarCount.FindControl("eOtherCount_OB_Edit");
            double vOtherCount_OB = 0.0;
            double.TryParse(eOtherCount_OB.Text.Trim(), out vOtherCount_OB);
            double vRuleCount_EA = 0.0;
            TextBox eRuleCount_EA = (TextBox)fvCarCount.FindControl("eRuleCount_EA_Edit");
            double.TryParse(eRuleCount_EA.Text.Trim(), out vRuleCount_EA);
            TextBox eHighWayCount_EA = (TextBox)fvCarCount.FindControl("eHighWayCount_EA_Edit");
            double vHighWayCount_EA = 0.0;
            double.TryParse(eHighWayCount_EA.Text.Trim(), out vHighWayCount_EA);
            TextBox eBusCount_EA = (TextBox)fvCarCount.FindControl("eBusCount_EA_Edit");
            double vBusCount_EA = 0.0;
            double.TryParse(eBusCount_EA.Text.Trim(), out vBusCount_EA);
            TextBox eTourCount_EA = (TextBox)fvCarCount.FindControl("eTourCount_EA_Edit");
            double vTourCount_EA = 0.0;
            double.TryParse(eTourCount_EA.Text.Trim(), out vTourCount_EA);
            TextBox eOtherCount_EA = (TextBox)fvCarCount.FindControl("eOtherCount_EA_Edit");
            double vOtherCount_EA = 0.0;
            double.TryParse(eOtherCount_EA.Text.Trim(), out vOtherCount_EA);
            TextBox eRuleCount_EB = (TextBox)fvCarCount.FindControl("eRuleCount_EB_Edit");
            double vRuleCount_EB = 0.0;
            double.TryParse(eRuleCount_EB.Text.Trim(), out vRuleCount_EB);
            TextBox eHighWayCount_EB = (TextBox)fvCarCount.FindControl("eHighWayCount_EB_Edit");
            double vHighWayCount_EB = 0.0;
            double.TryParse(eHighWayCount_EB.Text.Trim(), out vHighWayCount_EB);
            TextBox eBusCount_EB = (TextBox)fvCarCount.FindControl("eBusCount_EB_Edit");
            double vBusCount_EB = 0.0;
            double.TryParse(eBusCount_EB.Text.Trim(), out vBusCount_EB);
            TextBox eTourCount_EB = (TextBox)fvCarCount.FindControl("eTourCount_EB_Edit");
            double vTourCount_EB = 0.0;
            double.TryParse(eTourCount_EB.Text.Trim(), out vTourCount_EB);
            TextBox eOtherCount_EB = (TextBox)fvCarCount.FindControl("eOtherCount_EB_Edit");
            double vOtherCount_EB = 0.0;
            double.TryParse(eOtherCount_EB.Text.Trim(), out vOtherCount_EB);
            TextBox eCarCount = (TextBox)fvCarCount.FindControl("eCarCount_Edit");

            Label eRuleCount_OT = (Label)fvCarCount.FindControl("eRuleCount_OT_Edit");
            Label eHighWayCount_OT = (Label)fvCarCount.FindControl("eHighWayCount_OT_Edit");
            Label eBusCount_OT = (Label)fvCarCount.FindControl("eBusCount_OT_Edit");
            Label eTourCount_OT = (Label)fvCarCount.FindControl("eTourCount_OT_Edit");
            Label eOtherCount_OT = (Label)fvCarCount.FindControl("eOtherCount_OT_Edit");
            Label eRuleCount_ET = (Label)fvCarCount.FindControl("eRuleCount_ET_Edit");
            Label eHighWayCount_ET = (Label)fvCarCount.FindControl("eHighWayCount_ET_Edit");
            Label eBusCount_ET = (Label)fvCarCount.FindControl("eBusCount_ET_Edit");
            Label eTourCount_ET = (Label)fvCarCount.FindControl("eTourCount_ET_Edit");
            Label eOtherCount_ET = (Label)fvCarCount.FindControl("eOtherCount_ET_Edit");
            Label eCarSumA = (Label)fvCarCount.FindControl("eCarSumA_Edit");
            Label eCarSumB = (Label)fvCarCount.FindControl("eCarSumB_Edit");
            Label eCarSum_OT = (Label)fvCarCount.FindControl("eCarSum_OT_Edit");
            Label eCarSum_ET = (Label)fvCarCount.FindControl("eCarSum_ET_Edit");
            Label eCarSum_OA = (Label)fvCarCount.FindControl("eCarSum_OA_Edit");
            Label eCarSum_EA = (Label)fvCarCount.FindControl("eCarSum_EA_Edit");
            Label eCarSum_OB = (Label)fvCarCount.FindControl("eCarSum_OB_Edit");
            Label eCarSum_EB = (Label)fvCarCount.FindControl("eCarSum_EB_Edit");

            eRuleCount_OT.Text = (vRuleCount_OA + vRuleCount_OB).ToString();
            eRuleCount_ET.Text = (vRuleCount_EA + vRuleCount_EB).ToString();
            eHighWayCount_OT.Text = (vHighWayCount_OA + vHighWayCount_OB).ToString();
            eHighWayCount_ET.Text = (vHighWayCount_EA + vHighWayCount_EB).ToString();
            eBusCount_OT.Text = (vBusCount_OA + vBusCount_OB).ToString();
            eBusCount_ET.Text = (vBusCount_EA + vBusCount_EB).ToString();
            eTourCount_OT.Text = (vTourCount_OA + vTourCount_OB).ToString();
            eTourCount_ET.Text = (vTourCount_EA + vTourCount_EB).ToString();
            eOtherCount_OT.Text = (vOtherCount_OA + vOtherCount_OB).ToString();
            eOtherCount_ET.Text = (vOtherCount_EA + vOtherCount_EB).ToString();
            eCarSum_OA.Text = (vRuleCount_OA + vHighWayCount_OA + vBusCount_OA + vTourCount_OA + vOtherCount_OA).ToString();
            eCarSum_OB.Text = (vRuleCount_OB + vHighWayCount_OB + vBusCount_OB + vTourCount_OB + vOtherCount_OB).ToString();
            eCarSum_EA.Text = (vRuleCount_EA + vHighWayCount_EA + vBusCount_EA + vTourCount_EA + vOtherCount_EA).ToString();
            eCarSum_EB.Text = (vRuleCount_EB + vHighWayCount_EB + vBusCount_EB + vTourCount_EB + vOtherCount_EB).ToString();
            eCarSum_OT.Text = (vRuleCount_OA + vHighWayCount_OA + vBusCount_OA + vTourCount_OA + vOtherCount_OA + vRuleCount_OB + vHighWayCount_OB + vBusCount_OB + vTourCount_OB + vOtherCount_OB).ToString();
            eCarSum_ET.Text = (vRuleCount_EA + vHighWayCount_EA + vBusCount_EA + vTourCount_EA + vOtherCount_EA + vRuleCount_EB + vHighWayCount_EB + vBusCount_EB + vTourCount_EB + vOtherCount_EB).ToString();
            eCarSumA.Text = (vRuleCount_OA + vHighWayCount_OA + vBusCount_OA + vTourCount_OA + vOtherCount_OA + vRuleCount_EA + vHighWayCount_EA + vBusCount_EA + vTourCount_EA + vOtherCount_EA).ToString();
            eCarSumB.Text = (vRuleCount_OB + vHighWayCount_OB + vBusCount_OB + vTourCount_OB + vOtherCount_OB + vRuleCount_EB + vHighWayCount_EB + vBusCount_EB + vTourCount_EB + vOtherCount_EB).ToString();
            eCarCount.Text = (vRuleCount_OA + vHighWayCount_OA + vBusCount_OA + vTourCount_OA + vOtherCount_OA + vRuleCount_EA + vHighWayCount_EA + vBusCount_EA + vTourCount_EA + vOtherCount_EA + vRuleCount_OB + vHighWayCount_OB + vBusCount_OB + vTourCount_OB + vOtherCount_OB + vRuleCount_EB + vHighWayCount_EB + vBusCount_EB + vTourCount_EB + vOtherCount_EB).ToString();
        }

        protected void bbEdit_List_Click(object sender, EventArgs e)
        {
            fvCarCount.ChangeMode(FormViewMode.Edit);
        }

        protected void UpdateButton_Click(object sender, EventArgs e)
        {
            string vUpdateStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eDepNoYM = (Label)fvCarCount.FindControl("eDepNoYM_Edit");
            if (eDepNoYM != null)
            {
                vUpdateStr = "UPDATE CarCount " + Environment.NewLine +
                             "   SET ModifyYM = GetDate(), ModifyMan = @ModifyMan, CarCount = @CarCount, DriveRange = @DriveRange, Remark = @Remark, " + Environment.NewLine +
                             "       KMS_BUS = @KMS_BUS, KMS_Rent = @KMS_Rent, KMS_Tour = @KMS_Tour, KMS_Trans = @KMS_Trans, KMS_Spec = @KMS_Spec, " + Environment.NewLine +
                             "       RuleCount_OA = @RuleCount_OA, RuleCount_EA = @RuleCount_EA, RuleCount_OB = @RuleCount_OB, RuleCount_EB = @RuleCount_EB, RuleCount_OT = @RuleCount_OT, RuleCount_ET = @RuleCount_ET, " + Environment.NewLine +
                             "       BusCount_OA = @BusCount_OA, BusCount_EA = @BusCount_EA, BusCount_OB = @BusCount_OB, BusCount_EB = @BusCount_EB, BusCount_OT = @BusCount_OT, BusCount_ET = @BusCount_ET, " + Environment.NewLine +
                             "       TourCount_OA = @TourCount_OA, TourCount_EA = @TourCount_EA, TourCount_OB = @TourCount_OB, TourCount_EB = @TourCount_EB, TourCount_OT = @TourCount_OT, TourCount_ET = @TourCount_ET, " + Environment.NewLine +
                             "       CarSum_OA = @CarSum_OA, CarSum_EA = @CarSum_EA, CarSum_OB = @CarSum_OB, CarSum_EB = @CarSum_EB, CarSum_OT = @CarSum_OT, CarSum_ET = @CarSum_ET, " + Environment.NewLine +
                             "       RuleKMS_OA = @RuleKMS_OA, RuleKMS_EA = @RuleKMS_EA, RuleKMS_OB = @RuleKMS_OB, RuleKMS_EB = @RuleKMS_EB, RuleKMS_OT = @RuleKMS_OT, RuleKMS_ET = @RuleKMS_ET, " + Environment.NewLine +
                             "       BusKMS_OA = @BusKMS_OA, BusKMS_EA = @BusKMS_EA, BusKMS_OB = @BusKMS_OB, BusKMS_EB = @BusKMS_EB, BusKMS_OT = @BusKMS_OT, BusKMS_ET = @BusKMS_ET, " + Environment.NewLine +
                             "       SpecKMS_OA = @SpecKMS_OA, SpecKMS_EA = @SpecKMS_EA, SpecKMS_OB = @SpecKMS_OB, SpecKMS_EB = @SpecKMS_EB, SpecKMS_OT = @SpecKMS_OT, SpecKMS_ET = @SpecKMS_ET, " + Environment.NewLine +
                             "       RentKMS_OA = @RentKMS_OA, RentKMS_EA = @RentKMS_EA, RentKMS_OB = @RentKMS_OB, RentKMS_EB = @RentKMS_EB, RentKMS_OT = @RentKMS_OT, RentKMS_ET = @RentKMS_ET, " + Environment.NewLine +
                             "       TransKMS_OA = @TransKMS_OA, TransKMS_EA = @TransKMS_EA, TransKMS_OB = @TransKMS_OB, TransKMS_EB = @TransKMS_EB, TransKMS_OT = @TransKMS_OT, TransKMS_ET = @TransKMS_ET, " + Environment.NewLine +
                             "       TourKMS_OA = @TourKMS_OA, TourKMS_EA = @TourKMS_EA, TourKMS_OB = @TourKMS_OB, TourKMS_EB = @TourKMS_EB, TourKMS_OT = @TourKMS_OT, TourKMS_ET = @TourKMS_ET, " + Environment.NewLine +
                             "       NoneBusiKMS_OA = @NoneBusiKMS_OA, NoneBusiKMS_EA = @NoneBusiKMS_EA, NoneBusiKMS_OB = @NoneBusiKMS_OB, NoneBusiKMS_EB = @NoneBusiKMS_EB, NoneBusiKMS_OT = @NoneBusiKMS_OT, NoneBusiKMS_ET = @NoneBusiKMS_ET, " + Environment.NewLine +
                             "       TotalKMS_OA = @TotalKMS_OA, TotalKMS_EA = @TotalKMS_EA, TotalKMS_OB = @TotalKMS_OB, TotalKMS_EB = @TotalKMS_EB, TotalKMS_OT = @TotalKMS_OT, TotalKMS_ET = @TotalKMS_ET, " + Environment.NewLine +
                             "       HighWayCount_OA = @HighWayCount_OA, HighWayCount_EA = @HighWayCount_EA, HighWayCount_OB = @HighWayCount_OB, HighWayCount_EB = @HighWayCount_EB, HighWayCount_OT = @HighWayCount_OT, HighWayCount_ET = @HighWayCount_ET, " + Environment.NewLine +
                             "       OtherCount_OA = @OtherCount_OA, OtherCount_EA = @OtherCount_EA, OtherCount_OB = @OtherCount_OB, OtherCount_EB = @OtherCount_EB, OtherCount_OT = @OtherCount_OT, OtherCount_ET = @OtherCount_ET, " + Environment.NewLine +
                             "       KMS_Rule = @KMS_Rule, KMS_NoneBusi = @KMS_NoneBusi, KMS_Total = @KMS_Total, ModifyKMS = @ModifyKMS, SpecIncome = @SpecIncome, UsedCars = @UsedCars, ElecCardCount = @ElecCardCount, " + Environment.NewLine +
                             "       EmpCount = @EmpCount, DriverOD = @DriverOD, EmpCount_M = @EmpCount_M, EmpCount_F = @EmpCount_F, EmpCount_D = @EmpCount_D, BreakDown = @BreakDown " + Environment.NewLine +
                             " WHERE DepNoYM = @DepNoYM ";
                using (SqlDataSource dsUpdate = new SqlDataSource())
                {
                    dsUpdate.ConnectionString = vConnStr;
                    try
                    {
                        dsUpdate.UpdateCommand = vUpdateStr;
                        dsUpdate.UpdateParameters.Clear();
                        dsUpdate.UpdateParameters.Add(new Parameter("DepNoYM", DbType.String, eDepNoYM.Text.Trim()));
                        dsUpdate.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                        TextBox eCarCount = (TextBox)fvCarCount.FindControl("eCarCount_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("CarCount", DbType.Double, eCarCount.Text.Trim()));
                        Label eDriveRange = (Label)fvCarCount.FindControl("eDriveRange_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("DriveRange", DbType.Double, eDriveRange.Text.Trim()));
                        TextBox eRemark = (TextBox)fvCarCount.FindControl("eRemark_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("Remark", DbType.String, eRemark.Text.Trim()));
                        Label eKMS_Bus = (Label)fvCarCount.FindControl("eKMS_Bus_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("KMS_BUS", DbType.Double, eKMS_Bus.Text.Trim()));
                        Label eKMS_Rent = (Label)fvCarCount.FindControl("eKMS_Rent_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("KMS_Rent", DbType.Double, eKMS_Rent.Text.Trim()));
                        Label eKMS_Tour = (Label)fvCarCount.FindControl("eKMS_Tour_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("KMS_Tour", DbType.Double, eKMS_Tour.Text.Trim()));
                        Label eKMS_Trans = (Label)fvCarCount.FindControl("eKMS_Trans_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("KMS_Trans", DbType.Double, eKMS_Trans.Text.Trim()));
                        Label eKMS_Spec = (Label)fvCarCount.FindControl("eKMS_Spec_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("KMS_Spec", DbType.Double, eKMS_Spec.Text.Trim()));
                        TextBox eRuleCount_OA = (TextBox)fvCarCount.FindControl("eRuleCount_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RuleCount_OA", DbType.Double, eRuleCount_OA.Text.Trim()));
                        TextBox eRuleCount_EA = (TextBox)fvCarCount.FindControl("eRuleCount_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RuleCount_EA", DbType.Double, eRuleCount_EA.Text.Trim()));
                        TextBox eRuleCount_OB = (TextBox)fvCarCount.FindControl("eRuleCount_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RuleCount_OB", DbType.Double, eRuleCount_OB.Text.Trim()));
                        TextBox eRuleCount_EB = (TextBox)fvCarCount.FindControl("eRuleCount_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RuleCount_EB", DbType.Double, eRuleCount_EB.Text.Trim()));
                        Label eRuleCount_OT = (Label)fvCarCount.FindControl("eRuleCount_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RuleCount_OT", DbType.Double, eRuleCount_OT.Text.Trim()));
                        Label eRuleCount_ET = (Label)fvCarCount.FindControl("eRuleCount_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RuleCount_ET", DbType.Double, eRuleCount_ET.Text.Trim()));
                        TextBox eBusCount_OA = (TextBox)fvCarCount.FindControl("eBusCount_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("BusCount_OA", DbType.Double, eBusCount_OA.Text.Trim()));
                        TextBox eBusCount_EA = (TextBox)fvCarCount.FindControl("eBusCount_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("BusCount_EA", DbType.Double, eBusCount_EA.Text.Trim()));
                        TextBox eBusCount_OB = (TextBox)fvCarCount.FindControl("eBusCount_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("BusCount_OB", DbType.Double, eBusCount_OB.Text.Trim()));
                        TextBox eBusCount_EB = (TextBox)fvCarCount.FindControl("eBusCount_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("BusCount_EB", DbType.Double, eBusCount_EB.Text.Trim()));
                        Label eBusCount_OT = (Label)fvCarCount.FindControl("eBusCount_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("BusCount_OT", DbType.Double, eBusCount_OT.Text.Trim()));
                        Label eBusCount_ET = (Label)fvCarCount.FindControl("eBusCount_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("BusCount_ET", DbType.Double, eBusCount_ET.Text.Trim()));
                        TextBox eTourCount_OA = (TextBox)fvCarCount.FindControl("eTourCount_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TourCount_OA", DbType.Double, eTourCount_OA.Text.Trim()));
                        TextBox eTourCount_EA = (TextBox)fvCarCount.FindControl("eTourCount_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TourCount_EA", DbType.Double, eTourCount_EA.Text.Trim()));
                        TextBox eTourCount_OB = (TextBox)fvCarCount.FindControl("eTourCount_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TourCount_OB", DbType.Double, eTourCount_OB.Text.Trim()));
                        TextBox eTourCount_EB = (TextBox)fvCarCount.FindControl("eTourCount_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TourCount_EB", DbType.Double, eTourCount_EB.Text.Trim()));
                        Label eTourCount_OT = (Label)fvCarCount.FindControl("eTourCount_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TourCount_OT", DbType.Double, eTourCount_OT.Text.Trim()));
                        Label eTourCount_ET = (Label)fvCarCount.FindControl("eTourCount_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TourCount_ET", DbType.Double, eTourCount_ET.Text.Trim()));
                        Label eCarSum_OA = (Label)fvCarCount.FindControl("eCarSum_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("CarSum_OA", DbType.Double, eCarSum_OA.Text.Trim()));
                        Label eCarSum_EA = (Label)fvCarCount.FindControl("eCarSum_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("CarSum_EA", DbType.Double, eCarSum_EA.Text.Trim()));
                        Label eCarSum_OB = (Label)fvCarCount.FindControl("eCarSum_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("CarSum_OB", DbType.Double, eCarSum_OB.Text.Trim()));
                        Label eCarSum_EB = (Label)fvCarCount.FindControl("eCarSum_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("CarSum_EB", DbType.Double, eCarSum_EB.Text.Trim()));
                        Label eCarSum_OT = (Label)fvCarCount.FindControl("eCarSum_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("CarSum_OT", DbType.Double, eCarSum_OT.Text.Trim()));
                        Label eCarSum_ET = (Label)fvCarCount.FindControl("eCarSum_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("CarSum_ET", DbType.Double, eCarSum_ET.Text.Trim()));
                        TextBox eRuleKMS_OA = (TextBox)fvCarCount.FindControl("eRuleKMS_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RuleKMS_OA", DbType.Double, eRuleKMS_OA.Text.Trim()));
                        TextBox eRuleKMS_EA = (TextBox)fvCarCount.FindControl("eRuleKMS_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RuleKMS_EA", DbType.Double, eRuleKMS_EA.Text.Trim()));
                        TextBox eRuleKMS_OB = (TextBox)fvCarCount.FindControl("eRuleKMS_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RuleKMS_OB", DbType.Double, eRuleKMS_OB.Text.Trim()));
                        TextBox eRuleKMS_EB = (TextBox)fvCarCount.FindControl("eRuleKMS_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RuleKMS_EB", DbType.Double, eRuleKMS_EB.Text.Trim()));
                        Label eRuleKMS_OT = (Label)fvCarCount.FindControl("eRuleKMS_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RuleKMS_OT", DbType.Double, eRuleKMS_OT.Text.Trim()));
                        Label eRuleKMS_ET = (Label)fvCarCount.FindControl("eRuleKMS_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RuleKMS_ET", DbType.Double, eRuleKMS_ET.Text.Trim()));
                        TextBox eBusKMS_OA = (TextBox)fvCarCount.FindControl("eBusKMS_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("BusKMS_OA", DbType.Double, eBusKMS_OA.Text.Trim()));
                        TextBox eBusKMS_EA = (TextBox)fvCarCount.FindControl("eBusKMS_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("BusKMS_EA", DbType.Double, eBusKMS_EA.Text.Trim()));
                        TextBox eBusKMS_OB = (TextBox)fvCarCount.FindControl("eBusKMS_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("BusKMS_OB", DbType.Double, eBusKMS_OB.Text.Trim()));
                        TextBox eBusKMS_EB = (TextBox)fvCarCount.FindControl("eBusKMS_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("BusKMS_EB", DbType.Double, eBusKMS_EB.Text.Trim()));
                        Label eBusKMS_OT = (Label)fvCarCount.FindControl("eBusKMS_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("BusKMS_OT", DbType.Double, eBusKMS_OT.Text.Trim()));
                        Label eBusKMS_ET = (Label)fvCarCount.FindControl("eBusKMS_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("BusKMS_ET", DbType.Double, eBusKMS_ET.Text.Trim()));
                        TextBox eSpecKMS_OA = (TextBox)fvCarCount.FindControl("eSpecKMS_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("SpecKMS_OA", DbType.Double, eSpecKMS_OA.Text.Trim()));
                        TextBox eSpecKMS_EA = (TextBox)fvCarCount.FindControl("eSpecKMS_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("SpecKMS_EA", DbType.Double, eSpecKMS_EA.Text.Trim()));
                        TextBox eSpecKMS_OB = (TextBox)fvCarCount.FindControl("eSpecKMS_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("SpecKMS_OB", DbType.Double, eSpecKMS_OB.Text.Trim()));
                        TextBox eSpecKMS_EB = (TextBox)fvCarCount.FindControl("eSpecKMS_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("SpecKMS_EB", DbType.Double, eSpecKMS_EB.Text.Trim()));
                        Label eSpecKMS_OT = (Label)fvCarCount.FindControl("eSpecKMS_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("SpecKMS_OT", DbType.Double, eSpecKMS_OT.Text.Trim()));
                        Label eSpecKMS_ET = (Label)fvCarCount.FindControl("eSpecKMS_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("SpecKMS_ET", DbType.Double, eSpecKMS_ET.Text.Trim()));
                        TextBox eRentKMS_OA = (TextBox)fvCarCount.FindControl("eRentKMS_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RentKMS_OA", DbType.Double, eRentKMS_OA.Text.Trim()));
                        TextBox eRentKMS_EA = (TextBox)fvCarCount.FindControl("eRentKMS_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RentKMS_EA", DbType.Double, eRentKMS_EA.Text.Trim()));
                        TextBox eRentKMS_OB = (TextBox)fvCarCount.FindControl("eRentKMS_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RentKMS_OB", DbType.Double, eRentKMS_OB.Text.Trim()));
                        TextBox eRentKMS_EB = (TextBox)fvCarCount.FindControl("eRentKMS_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RentKMS_EB", DbType.Double, eRentKMS_EB.Text.Trim()));
                        Label eRentKMS_OT = (Label)fvCarCount.FindControl("eRentKMS_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RentKMS_OT", DbType.Double, eRentKMS_OT.Text.Trim()));
                        Label eRentKMS_ET = (Label)fvCarCount.FindControl("eRentKMS_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("RentKMS_ET", DbType.Double, eRentKMS_ET.Text.Trim()));
                        TextBox eTransKMS_OA = (TextBox)fvCarCount.FindControl("eTransKMS_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TransKMS_OA", DbType.Double, eTransKMS_OA.Text.Trim()));
                        TextBox eTransKMS_EA = (TextBox)fvCarCount.FindControl("eTransKMS_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TransKMS_EA", DbType.Double, eTransKMS_EA.Text.Trim()));
                        TextBox eTransKMS_OB = (TextBox)fvCarCount.FindControl("eTransKMS_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TransKMS_OB", DbType.Double, eTransKMS_OB.Text.Trim()));
                        TextBox eTransKMS_EB = (TextBox)fvCarCount.FindControl("eTransKMS_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TransKMS_EB", DbType.Double, eTransKMS_EB.Text.Trim()));
                        Label eTransKMS_OT = (Label)fvCarCount.FindControl("eTransKMS_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TransKMS_OT", DbType.Double, eTransKMS_OT.Text.Trim()));
                        Label eTransKMS_ET = (Label)fvCarCount.FindControl("eTransKMS_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TransKMS_ET", DbType.Double, eTransKMS_ET.Text.Trim()));
                        TextBox eTourKMS_OA = (TextBox)fvCarCount.FindControl("eTourKMS_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TourKMS_OA", DbType.Double, eTourKMS_OA.Text.Trim()));
                        TextBox eTourKMS_EA = (TextBox)fvCarCount.FindControl("eTourKMS_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TourKMS_EA", DbType.Double, eTourKMS_EA.Text.Trim()));
                        TextBox eTourKMS_OB = (TextBox)fvCarCount.FindControl("eTourKMS_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TourKMS_OB", DbType.Double, eTourKMS_OB.Text.Trim()));
                        TextBox eTourKMS_EB = (TextBox)fvCarCount.FindControl("eTourKMS_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TourKMS_EB", DbType.Double, eTourKMS_EB.Text.Trim()));
                        Label eTourKMS_OT = (Label)fvCarCount.FindControl("eTourKMS_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TourKMS_OT", DbType.Double, eTourKMS_OT.Text.Trim()));
                        Label eTourKMS_ET = (Label)fvCarCount.FindControl("eTourKMS_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TourKMS_ET", DbType.Double, eTourKMS_ET.Text.Trim()));
                        TextBox eNoneBusiKMS_OA = (TextBox)fvCarCount.FindControl("eNoneBusiKMS_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("NoneBusiKMS_OA", DbType.Double, eNoneBusiKMS_OA.Text.Trim()));
                        TextBox eNoneBusiKMS_EA = (TextBox)fvCarCount.FindControl("eNoneBusiKMS_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("NoneBusiKMS_EA", DbType.Double, eNoneBusiKMS_EA.Text.Trim()));
                        TextBox eNoneBusiKMS_OB = (TextBox)fvCarCount.FindControl("eNoneBusiKMS_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("NoneBusiKMS_OB", DbType.Double, eNoneBusiKMS_OB.Text.Trim()));
                        TextBox eNoneBusiKMS_EB = (TextBox)fvCarCount.FindControl("eNoneBusiKMS_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("NoneBusiKMS_EB", DbType.Double, eNoneBusiKMS_EB.Text.Trim()));
                        Label eNoneBusiKMS_OT = (Label)fvCarCount.FindControl("eNoneBusiKMS_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("NoneBusiKMS_OT", DbType.Double, eNoneBusiKMS_OT.Text.Trim()));
                        Label eNoneBusiKMS_ET = (Label)fvCarCount.FindControl("eNoneBusiKMS_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("NoneBusiKMS_ET", DbType.Double, eNoneBusiKMS_ET.Text.Trim()));
                        Label eTotalKMS_OA = (Label)fvCarCount.FindControl("eTotalKMS_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TotalKMS_OA", DbType.Double, eTotalKMS_OA.Text.Trim()));
                        Label eTotalKMS_EA = (Label)fvCarCount.FindControl("eTotalKMS_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TotalKMS_EA", DbType.Double, eTotalKMS_EA.Text.Trim()));
                        Label eTotalKMS_OB = (Label)fvCarCount.FindControl("eTotalKMS_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TotalKMS_OB", DbType.Double, eTotalKMS_OB.Text.Trim()));
                        Label eTotalKMS_EB = (Label)fvCarCount.FindControl("eTotalKMS_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TotalKMS_EB", DbType.Double, eTotalKMS_EB.Text.Trim()));
                        Label eTotalKMS_OT = (Label)fvCarCount.FindControl("eTotalKMS_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TotalKMS_OT", DbType.Double, eTotalKMS_OT.Text.Trim()));
                        Label eTotalKMS_ET = (Label)fvCarCount.FindControl("eTotalKMS_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("TotalKMS_ET", DbType.Double, eTotalKMS_ET.Text.Trim()));
                        TextBox eHighWayCount_OA = (TextBox)fvCarCount.FindControl("eHighWayCount_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("HighWayCount_OA", DbType.Double, eHighWayCount_OA.Text.Trim()));
                        TextBox eHighWayCount_EA = (TextBox)fvCarCount.FindControl("eHighWayCount_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("HighWayCount_EA", DbType.Double, eHighWayCount_EA.Text.Trim()));
                        TextBox eHighWayCount_OB = (TextBox)fvCarCount.FindControl("eHighWayCount_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("HighWayCount_OB", DbType.Double, eHighWayCount_OB.Text.Trim()));
                        TextBox eHighWayCount_EB = (TextBox)fvCarCount.FindControl("eHighWayCount_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("HighWayCount_EB", DbType.Double, eHighWayCount_EB.Text.Trim()));
                        Label eHighWayCount_OT = (Label)fvCarCount.FindControl("eHighWayCount_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("HighWayCount_OT", DbType.Double, eHighWayCount_OT.Text.Trim()));
                        Label eHighWayCount_ET = (Label)fvCarCount.FindControl("eHighWayCount_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("HighWayCount_ET", DbType.Double, eHighWayCount_ET.Text.Trim()));
                        TextBox eOtherCount_OA = (TextBox)fvCarCount.FindControl("eOtherCount_OA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("OtherCount_OA", DbType.Double, eOtherCount_OA.Text.Trim()));
                        TextBox eOtherCount_EA = (TextBox)fvCarCount.FindControl("eOtherCount_EA_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("OtherCount_EA", DbType.Double, eOtherCount_EA.Text.Trim()));
                        TextBox eOtherCount_OB = (TextBox)fvCarCount.FindControl("eOtherCount_OB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("OtherCount_OB", DbType.Double, eOtherCount_OB.Text.Trim()));
                        TextBox eOtherCount_EB = (TextBox)fvCarCount.FindControl("eOtherCount_EB_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("OtherCount_EB", DbType.Double, eOtherCount_EB.Text.Trim()));
                        Label eOtherCount_OT = (Label)fvCarCount.FindControl("eOtherCount_OT_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("OtherCount_OT", DbType.Double, eOtherCount_OT.Text.Trim()));
                        Label eOtherCount_ET = (Label)fvCarCount.FindControl("eOtherCount_ET_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("OtherCount_ET", DbType.Double, eOtherCount_ET.Text.Trim()));
                        Label eKMS_Rule = (Label)fvCarCount.FindControl("eKMS_Rule_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("KMS_Rule", DbType.Double, eKMS_Rule.Text.Trim()));
                        Label eKMS_NoneBusi = (Label)fvCarCount.FindControl("eKMS_NoneBusi_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("KMS_NoneBusi", DbType.Double, eKMS_NoneBusi.Text.Trim()));
                        Label eKMS_Total = (Label)fvCarCount.FindControl("eKMS_Total_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("KMS_Total", DbType.Double, eKMS_Total.Text.Trim()));
                        TextBox eModifyKMS = (TextBox)fvCarCount.FindControl("eModifyKMS_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("ModifyKMS", DbType.Double, eModifyKMS.Text.Trim()));
                        TextBox eSpecIncome = (TextBox)fvCarCount.FindControl("eSpecIncome_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("SpecIncome", DbType.Double, eSpecIncome.Text.Trim()));
                        TextBox eUsedCars = (TextBox)fvCarCount.FindControl("eUsedCars_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("UsedCars", DbType.Double, eUsedCars.Text.Trim()));
                        TextBox eElecCardCount = (TextBox)fvCarCount.FindControl("eElecCardCount_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("ElecCardCount", DbType.Double, eElecCardCount.Text.Trim()));
                        Label eEmpCount = (Label)fvCarCount.FindControl("eEmpCount_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("EmpCount", DbType.Double, eEmpCount.Text.Trim()));
                        TextBox eEmpCount_M = (TextBox)fvCarCount.FindControl("eEmpCount_M_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("EmpCount_M", DbType.Double, eEmpCount_M.Text.Trim()));
                        TextBox eEmpCount_F = (TextBox)fvCarCount.FindControl("eEmpCount_F_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("EmpCount_F", DbType.Double, eEmpCount_F.Text.Trim()));
                        TextBox eEmpCount_D = (TextBox)fvCarCount.FindControl("eEmpCount_D_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("EmpCount_D", DbType.Double, eEmpCount_D.Text.Trim()));
                        TextBox eDriverOD = (TextBox)fvCarCount.FindControl("eDriverOD_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("DriverOD", DbType.Double, eDriverOD.Text.Trim()));
                        TextBox eBreakDown = (TextBox)fvCarCount.FindControl("eBreakDown_Edit");
                        dsUpdate.UpdateParameters.Add(new Parameter("BreakDown", DbType.Double, eBreakDown.Text.Trim()));
                        dsUpdate.Update();
                        fvCarCount.ChangeMode(FormViewMode.ReadOnly);
                        dsCarCount.DataBind();
                        fvCarCount.DataBind();
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

        protected void eEmpCount_M_Edit_TextChanged(object sender, EventArgs e)
        {
            Label eEmpCount = (Label)fvCarCount.FindControl("eEmpCount_Edit");
            TextBox eEmpCount_M = (TextBox)fvCarCount.FindControl("eEmpCount_M_Edit");
            double vEmpCount_M = 0.0;
            double.TryParse(eEmpCount_M.Text.Trim(), out vEmpCount_M);
            TextBox eEmpCount_F = (TextBox)fvCarCount.FindControl("eEmpCount_F_Edit");
            double vEmpCount_F = 0.0;
            double.TryParse(eEmpCount_F.Text.Trim(), out vEmpCount_F);
            TextBox eEmpCount_D = (TextBox)fvCarCount.FindControl("eEmpCount_D_Edit");
            double vEmpCount_D = 0.0;
            double.TryParse(eEmpCount_D.Text.Trim(), out vEmpCount_D);

            eEmpCount.Text = (vEmpCount_M + vEmpCount_F + vEmpCount_D).ToString();
        }

        protected void eRuleKMS_OA_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eRuleKMS_OA = (TextBox)fvCarCount.FindControl("eRuleKMS_OA_Edit");
            double vRuleKMS_OA = 0.0;
            double.TryParse(eRuleKMS_OA.Text.Trim(), out vRuleKMS_OA);
            TextBox eRuleKMS_EA = (TextBox)fvCarCount.FindControl("eRuleKMS_EA_Edit");
            double vRuleKMS_EA = 0.0;
            double.TryParse(eRuleKMS_EA.Text.Trim(), out vRuleKMS_EA);
            TextBox eRuleKMS_OB = (TextBox)fvCarCount.FindControl("eRuleKMS_OB_Edit");
            double vRuleKMS_OB = 0.0;
            double.TryParse(eRuleKMS_OB.Text.Trim(), out vRuleKMS_OB);
            TextBox eRuleKMS_EB = (TextBox)fvCarCount.FindControl("eRuleKMS_EB_Edit");
            double vRuleKMS_EB = 0.0;
            double.TryParse(eRuleKMS_EB.Text.Trim(), out vRuleKMS_EB);

            TextBox eBusKMS_OA = (TextBox)fvCarCount.FindControl("eBusKMS_OA_Edit");
            double vBusKMS_OA = 0.0;
            double.TryParse(eBusKMS_OA.Text.Trim(), out vBusKMS_OA);
            TextBox eBusKMS_EA = (TextBox)fvCarCount.FindControl("eBusKMS_EA_Edit");
            double vBusKMS_EA = 0.0;
            double.TryParse(eBusKMS_EA.Text.Trim(), out vBusKMS_EA);
            TextBox eBusKMS_OB = (TextBox)fvCarCount.FindControl("eBusKMS_OB_Edit");
            double vBusKMS_OB = 0.0;
            double.TryParse(eBusKMS_OB.Text.Trim(), out vBusKMS_OB);
            TextBox eBusKMS_EB = (TextBox)fvCarCount.FindControl("eBusKMS_EB_Edit");
            double vBusKMS_EB = 0.0;
            double.TryParse(eBusKMS_EB.Text.Trim(), out vBusKMS_EB);

            TextBox eSpecKMS_OA = (TextBox)fvCarCount.FindControl("eSpecKMS_OA_Edit");
            double vSpecKMS_OA = 0.0;
            double.TryParse(eSpecKMS_OA.Text.Trim(), out vSpecKMS_OA);
            TextBox eSpecKMS_EA = (TextBox)fvCarCount.FindControl("eSpecKMS_EA_Edit");
            double vSpecKMS_EA = 0.0;
            double.TryParse(eSpecKMS_EA.Text.Trim(), out vSpecKMS_EA);
            TextBox eSpecKMS_OB = (TextBox)fvCarCount.FindControl("eSpecKMS_OB_Edit");
            double vSpecKMS_OB = 0.0;
            double.TryParse(eSpecKMS_OB.Text.Trim(), out vSpecKMS_OB);
            TextBox eSpecKMS_EB = (TextBox)fvCarCount.FindControl("eSpecKMS_EB_Edit");
            double vSpecKMS_EB = 0.0;
            double.TryParse(eSpecKMS_EB.Text.Trim(), out vSpecKMS_EB);

            TextBox eRentKMS_OA = (TextBox)fvCarCount.FindControl("eRentKMS_OA_Edit");
            double vRentKMS_OA = 0.0;
            double.TryParse(eRentKMS_OA.Text.Trim(), out vRentKMS_OA);
            TextBox eRentKMS_EA = (TextBox)fvCarCount.FindControl("eRentKMS_EA_Edit");
            double vRentKMS_EA = 0.0;
            double.TryParse(eRentKMS_EA.Text.Trim(), out vRentKMS_EA);
            TextBox eRentKMS_OB = (TextBox)fvCarCount.FindControl("eRentKMS_OB_Edit");
            double vRentKMS_OB = 0.0;
            double.TryParse(eRentKMS_OB.Text.Trim(), out vRentKMS_OB);
            TextBox eRentKMS_EB = (TextBox)fvCarCount.FindControl("eRentKMS_EB_Edit");
            double vRentKMS_EB = 0.0;
            double.TryParse(eRentKMS_EB.Text.Trim(), out vRentKMS_EB);

            TextBox eTransKMS_OA = (TextBox)fvCarCount.FindControl("eTransKMS_OA_Edit");
            double vTransKMS_OA = 0.0;
            double.TryParse(eTransKMS_OA.Text.Trim(), out vTransKMS_OA);
            TextBox eTransKMS_EA = (TextBox)fvCarCount.FindControl("eTransKMS_EA_Edit");
            double vTransKMS_EA = 0.0;
            double.TryParse(eTransKMS_EA.Text.Trim(), out vTransKMS_EA);
            TextBox eTransKMS_OB = (TextBox)fvCarCount.FindControl("eTransKMS_OB_Edit");
            double vTransKMS_OB = 0.0;
            double.TryParse(eTransKMS_OB.Text.Trim(), out vTransKMS_OB);
            TextBox eTransKMS_EB = (TextBox)fvCarCount.FindControl("eTransKMS_EB_Edit");
            double vTransKMS_EB = 0.0;
            double.TryParse(eTransKMS_EB.Text.Trim(), out vTransKMS_EB);

            TextBox eTourKMS_OA = (TextBox)fvCarCount.FindControl("eTourKMS_OA_Edit");
            double vTourKMS_OA = 0.0;
            double.TryParse(eTourKMS_OA.Text.Trim(), out vTourKMS_OA);
            TextBox eTourKMS_EA = (TextBox)fvCarCount.FindControl("eTourKMS_EA_Edit");
            double vTourKMS_EA = 0.0;
            double.TryParse(eTourKMS_EA.Text.Trim(), out vTourKMS_EA);
            TextBox eTourKMS_OB = (TextBox)fvCarCount.FindControl("eTourKMS_OB_Edit");
            double vTourKMS_OB = 0.0;
            double.TryParse(eTourKMS_OB.Text.Trim(), out vTourKMS_OB);
            TextBox eTourKMS_EB = (TextBox)fvCarCount.FindControl("eTourKMS_EB_Edit");
            double vTourKMS_EB = 0.0;
            double.TryParse(eTourKMS_EB.Text.Trim(), out vTourKMS_EB);

            TextBox eNoneBusiKMS_OA = (TextBox)fvCarCount.FindControl("eNoneBusiKMS_OA_Edit");
            double vNoneBusiKMS_OA = 0.0;
            double.TryParse(eNoneBusiKMS_OA.Text.Trim(), out vNoneBusiKMS_OA);
            TextBox eNoneBusiKMS_EA = (TextBox)fvCarCount.FindControl("eNoneBusiKMS_EA_Edit");
            double vNoneBusiKMS_EA = 0.0;
            double.TryParse(eNoneBusiKMS_EA.Text.Trim(), out vNoneBusiKMS_EA);
            TextBox eNoneBusiKMS_OB = (TextBox)fvCarCount.FindControl("eNoneBusiKMS_OB_Edit");
            double vNoneBusiKMS_OB = 0.0;
            double.TryParse(eNoneBusiKMS_OB.Text.Trim(), out vNoneBusiKMS_OB);
            TextBox eNoneBusiKMS_EB = (TextBox)fvCarCount.FindControl("eNoneBusiKMS_EB_Edit");
            double vNoneBusiKMS_EB = 0.0;
            double.TryParse(eNoneBusiKMS_EB.Text.Trim(), out vNoneBusiKMS_EB);
            TextBox eModifyKMS = (TextBox)fvCarCount.FindControl("eModifyKMS_Edit");
            double vModifyKMS = 0.0;
            double.TryParse(eModifyKMS.Text.Trim(), out vModifyKMS);

            Label eRuleKMS_OT = (Label)fvCarCount.FindControl("eRuleKMS_OT_Edit");
            eRuleKMS_OT.Text = (vRuleKMS_OA + vRuleKMS_OB).ToString();
            Label eRuleKMS_ET = (Label)fvCarCount.FindControl("eRuleKMS_ET_Edit");
            eRuleKMS_ET.Text = (vRuleKMS_EA + vRuleKMS_EB).ToString();
            Label eBusKMS_OT = (Label)fvCarCount.FindControl("eBusKMS_OT_Edit");
            eBusKMS_OT.Text = (vBusKMS_OA + vBusKMS_OB).ToString();
            Label eBusKMS_ET = (Label)fvCarCount.FindControl("eBusKMS_ET_Edit");
            eBusKMS_ET.Text = (vBusKMS_EA + vBusKMS_EB).ToString();
            Label eSpecKMS_OT = (Label)fvCarCount.FindControl("eSpecKMS_OT_Edit");
            eSpecKMS_OT.Text = (vSpecKMS_OA + vSpecKMS_OB).ToString();
            Label eSpecKMS_ET = (Label)fvCarCount.FindControl("eSpecKMS_ET_Edit");
            eSpecKMS_ET.Text = (vSpecKMS_EA + vSpecKMS_EB).ToString();
            Label eRentKMS_OT = (Label)fvCarCount.FindControl("eRentKMS_OT_Edit");
            eRentKMS_OT.Text = (vRentKMS_OA + vRentKMS_OB).ToString();
            Label eRentKMS_ET = (Label)fvCarCount.FindControl("eRentKMS_ET_Edit");
            eRentKMS_ET.Text = (vRentKMS_EA + vRentKMS_EB).ToString();
            Label eTransKMS_OT = (Label)fvCarCount.FindControl("eTransKMS_OT_Edit");
            eTransKMS_OT.Text = (vTransKMS_OA + vTransKMS_OB).ToString();
            Label eTransKMS_ET = (Label)fvCarCount.FindControl("eTransKMS_ET_Edit");
            eTransKMS_ET.Text = (vTransKMS_EA + vTransKMS_EB).ToString();
            Label eTourKMS_OT = (Label)fvCarCount.FindControl("eTourKMS_OT_Edit");
            eTourKMS_OT.Text = (vTourKMS_OA + vTourKMS_OB).ToString();
            Label eTourKMS_ET = (Label)fvCarCount.FindControl("eTourKMS_ET_Edit");
            eTourKMS_ET.Text = (vTourKMS_EA + vTourKMS_EB).ToString();
            Label eNoneBusiKMS_OT = (Label)fvCarCount.FindControl("eNoneBusiKMS_OT_Edit");
            eNoneBusiKMS_OT.Text = (vNoneBusiKMS_OA + vNoneBusiKMS_OB).ToString();
            Label eNoneBusiKMS_ET = (Label)fvCarCount.FindControl("eNoneBusiKMS_ET_Edit");
            eNoneBusiKMS_ET.Text = (vNoneBusiKMS_EA + vNoneBusiKMS_EB).ToString();
            Label eTotalKMS_OA = (Label)fvCarCount.FindControl("eTotalKMS_OA_Edit");
            eTotalKMS_OA.Text = (vRuleKMS_OA + vBusKMS_OA + vSpecKMS_OA + vRentKMS_OA + vTransKMS_OA + vTourKMS_OA + vNoneBusiKMS_OA).ToString();
            Label eTotalKMS_EA = (Label)fvCarCount.FindControl("eTotalKMS_EA_Edit");
            eTotalKMS_EA.Text = (vRuleKMS_EA + vBusKMS_EA + vSpecKMS_EA + vRentKMS_EA + vTransKMS_EA + vTourKMS_EA + vNoneBusiKMS_EA).ToString();
            Label eTotalKMS_OB = (Label)fvCarCount.FindControl("eTotalKMS_OB_Edit");
            eTotalKMS_OB.Text = (vRuleKMS_OB + vBusKMS_OB + vSpecKMS_OB + vRentKMS_OB + vTransKMS_OB + vTourKMS_OB + vNoneBusiKMS_OB).ToString();
            Label eTotalKMS_EB = (Label)fvCarCount.FindControl("eTotalKMS_EB_Edit");
            eTotalKMS_EB.Text = (vRuleKMS_EB + vBusKMS_EB + vSpecKMS_EB + vRentKMS_EB + vTransKMS_EB + vTourKMS_EB + vNoneBusiKMS_EB).ToString();
            Label eTotalKMS_OT = (Label)fvCarCount.FindControl("eTotalKMS_OT_Edit");
            eTotalKMS_OA.Text = (vRuleKMS_OA + vBusKMS_OA + vSpecKMS_OA + vRentKMS_OA + vTransKMS_OA + vTourKMS_OA + vNoneBusiKMS_OA + vRuleKMS_OB + vBusKMS_OB + vSpecKMS_OB + vRentKMS_OB + vTransKMS_OB + vTourKMS_OB + vNoneBusiKMS_OB).ToString();
            Label eTotalKMS_ET = (Label)fvCarCount.FindControl("eTotalKMS_ET_Edit");
            eTotalKMS_EA.Text = (vRuleKMS_EA + vBusKMS_EA + vSpecKMS_EA + vRentKMS_EA + vTransKMS_EA + vTourKMS_EA + vNoneBusiKMS_EA + vRuleKMS_EB + vBusKMS_EB + vSpecKMS_EB + vRentKMS_EB + vTransKMS_EB + vTourKMS_EB + vNoneBusiKMS_EB).ToString();
            Label eKMS_Bus = (Label)fvCarCount.FindControl("eKMS_Bus_Edit");
            eKMS_Bus.Text = (vBusKMS_OA + vBusKMS_EA + vBusKMS_OB + vBusKMS_EB).ToString();
            Label eKMS_Rent = (Label)fvCarCount.FindControl("eKMS_Rent_Edit");
            eKMS_Rent.Text = (vRentKMS_OA + vRentKMS_EA + vRentKMS_OB + vRentKMS_EB).ToString();
            Label eKMS_Tour = (Label)fvCarCount.FindControl("eKMS_Tour_Edit");
            eKMS_Tour.Text = (vTourKMS_OA + vTourKMS_EA + vTourKMS_OB + vTourKMS_EB).ToString();
            Label eKMS_Trans = (Label)fvCarCount.FindControl("eKMS_Trans_Edit");
            eKMS_Trans.Text = (vTransKMS_OA + vTransKMS_EA + vTransKMS_OB + vTransKMS_EB).ToString();
            Label eKMS_Spec = (Label)fvCarCount.FindControl("eKMS_Spec_Edit");
            eKMS_Spec.Text = (vSpecKMS_OA + vSpecKMS_EA + vSpecKMS_OB + vSpecKMS_EB).ToString();
            Label eKMS_Rule = (Label)fvCarCount.FindControl("eKMS_Rule_Edit");
            eKMS_Rule.Text = (vRuleKMS_OA + vRuleKMS_EA + vRuleKMS_OB + vRuleKMS_EB).ToString();
            Label eKMS_NoneBusi = (Label)fvCarCount.FindControl("eKMS_NoneBusi_Edit");
            eKMS_NoneBusi.Text = (vNoneBusiKMS_OA + vNoneBusiKMS_EA + vNoneBusiKMS_OB + vNoneBusiKMS_EB).ToString();
            Label eKMS_Total = (Label)fvCarCount.FindControl("eKMS_Total_Edit");
            eKMS_Total.Text = (vRuleKMS_OA + vBusKMS_OA + vSpecKMS_OA + vRentKMS_OA + vTransKMS_OA + vTourKMS_OA + vNoneBusiKMS_OA + vRuleKMS_OB + vBusKMS_OB + vSpecKMS_OB + vRentKMS_OB + vTransKMS_OB + vTourKMS_OB + vNoneBusiKMS_OB + vRuleKMS_EA + vBusKMS_EA + vSpecKMS_EA + vRentKMS_EA + vTransKMS_EA + vTourKMS_EA + vNoneBusiKMS_EA + vRuleKMS_EB + vBusKMS_EB + vSpecKMS_EB + vRentKMS_EB + vTransKMS_EB + vTourKMS_EB + vNoneBusiKMS_EB).ToString();
        }
    }
}