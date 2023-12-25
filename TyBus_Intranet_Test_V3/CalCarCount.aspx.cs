using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Amaterasu_Function;
using System.Data.SqlClient;
using System.Globalization;

namespace TyBus_Intranet_Test_V3
{
    public partial class CalCarCount : System.Web.UI.Page
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
                    vSelectStr = "select cast('" + eCalYM_Temp.ToString("yyyy/MM/dd") + "' as DateTime) as CalYM, CompanyNo as DepNo,('" + eCalYM_Search.Text.Trim() + "' + CompanyNo) as DepNoYM " + Environment.NewLine +
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
                                 "  from Car_infoA group by CompanyNo ";
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
                                                    dtTarget.Rows[i]["HighWayCount_OA"] = (vHasOldData)? vTempFloatA : (vTempFloatA + vTempFloat);
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
                                            if (drModifyKMS["Class"].ToString().Trim()=="甲") //大巴
                                            {
                                                vTempSpecKMS = double.Parse(dtTarget.Rows[i]["SpecKMS_OA"].ToString().Trim());
                                                vTempBusKMS = double.Parse(dtTarget.Rows[i]["BusKMS_OA"].ToString().Trim());
                                                dtTarget.Rows[i]["SpecKMS_OA"] = vTempSpecKMS - vTempFloat;
                                                dtTarget.Rows[i]["BusKMS_OA"] = vTempBusKMS + vTempFloat;
                                            }
                                            else if (drModifyKMS["Class"].ToString().Trim()=="乙") //中巴
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
                    fvCarCount.DataBind();
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
    }
}