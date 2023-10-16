using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Globalization;

namespace TyBus_Intranet_Test_V3
{
    public partial class CarModifyList : System.Web.UI.Page
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
        private string vActionDateStr = "";

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
                    if (!IsPostBack)
                    {
                        eActionDate_Search.Text = (eActionDate_Search.Text.Trim() != "") ? eActionDate_Search.Text.Trim() : PF.GetChinsesDate(DateTime.Today);
                        plReport.Visible = false;
                    }
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    string vActionDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eActionDate_Search.ClientID;
                    string vActionDateScript = "window.open('" + vActionDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eActionDate_Search.Attributes["onClick"] = vActionDateScript;
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

        private string GetSelStr()
        {
            string vActionDate = (eActionDate_Search.Text.Trim() == "") ?
                               DateTime.Today.Year.ToString() + "/" + DateTime.Today.Month.ToString("D2") + "/" + DateTime.Today.Day.ToString("D2") :
                               DateTime.Parse(eActionDate_Search.Text.Trim()).Year.ToString() + "/" + DateTime.Parse(eActionDate_Search.Text.Trim()).Month.ToString("D2") + "/" + DateTime.Parse(eActionDate_Search.Text.Trim()).Day.ToString("D2");
            string vReturnStr = "select e.Car_ID as CarID, " + Environment.NewLine +
                                "       (select RTRIM(ClassTxt) from DBDICB where FKey = '車輛資料作業    Car_infoA       CAR_CLASS' and ClassNo = a.Car_Class)Car_Class_C, " + Environment.NewLine +
                                "       cast(Year(a.ProdDate) as varchar) + '.' + right('00' + cast(Month(a.ProdDate) as varchar), 2) ProdDate, " + Environment.NewLine +
                                "       a.SitQty as SiteQty, e.TranDate, " + Environment.NewLine +
                                "       (select[Name] from Department where DepNo = e.DepNo_O) DepNo_O_C, " + Environment.NewLine +
                                "       case when e.Tran_Type <> '1' then null else (select[Name] from Department where DepNo = e.DepNo) end DepNo_C, " + Environment.NewLine +
                                "       e.Tran_Type, " + Environment.NewLine +
                                "       (case when e.Tran_Type = '1' and e.DepNo_O = e.DepNo and(select count(CarTranNo) from Car_InfoE where Car_ID = e.Car_ID) = 1 then '新進車輛' " + Environment.NewLine +
                                "             when e.Tran_Type = '1' and(select count(CarTranNo) from Car_InfoE where Car_ID = e.Car_ID) <> 1 then '調站' " + Environment.NewLine +
                                "             else (select RTRIM(ClassTxt) from DBDICB where FKey = '車輛異動        Car_infoE       TRANTYPE' and ClassNo = e.Tran_Type) " + Environment.NewLine +
                                "              end) Tran_Type_C, e.Remark " + Environment.NewLine +
                                "  from Car_InfoE e left join Car_InfoA a on a.Car_ID = e.Car_ID " + Environment.NewLine +
                                //" where e.DecDate = '" + vActionDate + "' and e.PassDate is not null";
                                " where e.DecDate = '" + vActionDate + "' "; //2019.11.26 因應承辦單位要求不過濾是否過檔
            vActionDateStr = "民國" + (DateTime.Parse(eActionDate_Search.Text.Trim()).Year - 1911).ToString() + "年" +
                                      DateTime.Parse(eActionDate_Search.Text.Trim()).Month.ToString("D2") + "月" +
                                      DateTime.Parse(eActionDate_Search.Text.Trim()).Day.ToString("D2") + "日";
            return vReturnStr;
        }

        /// <summary>
        /// 預覽報表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPreview_Click(object sender, EventArgs e)
        {
            sdsCarModifyListP.SelectCommand = GetSelStr();
            //統計報表
            ReportDataSource rdsPrint = new ReportDataSource("CarModifyListP", sdsCarModifyListP);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\CarModifyListP.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "車輛調動通知單"));
            rvPrint.LocalReport.SetParameters(new ReportParameter("ActionDate", vActionDateStr));
            rvPrint.LocalReport.Refresh();
            plReport.Visible = true;
            plSearch.Visible = false;
            string vRecordActionDateStr = (eActionDate_Search.Text.Trim() != "") ? eActionDate_Search.Text.Trim() : "全部";
            string vRecordNote = "預覽報表_車輛調動通知單" + Environment.NewLine +
                                 "CarModifyList.aspx" + Environment.NewLine +
                                 "發文日期：" + vRecordActionDateStr;
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            rvPrint.LocalReport.DataSources.Clear();
            plReport.Visible = false;
            plSearch.Visible = true;
        }
    }
}