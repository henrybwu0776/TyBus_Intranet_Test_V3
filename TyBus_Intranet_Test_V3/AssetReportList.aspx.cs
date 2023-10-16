using Amaterasu_Function;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class AssetReportList : Page
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
                DateTime vToday = DateTime.Today;
                vLoginID = (Session["LoginID"] != null) ? Session["LoginID"].ToString().Trim() : "";
                vLoginName = (Session["LoginName"] != null) ? Session["LoginName"].ToString().Trim() : "";
                vLoginDepNo = (Session["LoginDepNo"] != null) ? Session["LoginDepNo"].ToString().Trim() : "";
                vLoginDepName = (Session["LoginDepName"] != null) ? Session["LoginDepName"].ToString().Trim() : "";
                vLoginTitle = (Session["LoginTitle"] != null) ? Session["LoginTitle"].ToString().Trim() : "";
                vLoginTitleName = (Session["LoginTitleName"] != null) ? Session["LoginTitleName"].ToString().Trim() : "";
                vLoginEmpType = (Session["LoginEmpType"] != null) ? Session["LoginEmpType"].ToString().Trim() : "";
                vComputerName = Environment.MachineName; //2021.09.27 新增取得電腦名稱

                if (vLoginID != "")
                {
                    if (!IsPostBack)
                    {
                        if (vToday.Month - 1 <= 0)
                        {
                            eStopDate_Year.Text = (eStopDate_Year.Text.Trim() == "") ? (vToday.Year - 1).ToString("D4") : eStopDate_Year.Text.Trim();
                            eStopDate_Month.Text = (eStopDate_Month.Text.Trim() == "") ? "12" : eStopDate_Month.Text.Trim();
                        }
                        else
                        {
                            eStopDate_Year.Text = (eStopDate_Year.Text.Trim() == "") ? vToday.Year.ToString("D4") : eStopDate_Year.Text.Trim();
                            eStopDate_Month.Text = (eStopDate_Month.Text.Trim() == "") ? ((vToday.Month) - 1).ToString("D2") : eStopDate_Month.Text.Trim();
                        }
                        plShowData.Visible = false;
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

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            plShowData.Visible = false;
            plAssetP001.Visible = false;
            plAssetP002.Visible = false;
            plAssetP003.Visible = false;
            plAssetP004.Visible = false;
            plAssetP005.Visible = false;
            plAssetP006.Visible = false;
            /* 預覽資料 */
            int vYear = (eStopDate_Year.Text.Trim() != "") ? Int32.Parse(eStopDate_Year.Text.Trim()) : DateTime.Today.Year;
            int vMonth = (eStopDate_Month.Text.Trim() != "") ? Int32.Parse(eStopDate_Month.Text.Trim()) : DateTime.Today.Month;
            int vNextMonth = (vMonth + 1 <= 12) ? vMonth + 1 : 1;
            int vNextYear = (vMonth + 1 <= 12) ? vYear : vYear + 1;
            //設定各種日期
            DateTime vEndOfLastYear = DateTime.Parse((vYear - 1).ToString("D4") + "/12/31"); //前一年的最後一天
            DateTime vFirstDayOfThisYear = DateTime.Parse(vYear.ToString("D4") + "/01/01"); //本年度的第一天
            DateTime vFirstDayOfThisMonth = DateTime.Parse(vYear.ToString("D4") + "/" + vMonth.ToString("D2") + "/01"); //當月第一天
            DateTime vFirstDayOfNextMonth = DateTime.Parse(vNextYear.ToString("D4") + "/" + vNextMonth.ToString("D2") + "/01"); //下個月第一天
            DateTime vLastDayOfThisMonth = DateTime.Parse(PF.GetMonthLastDay(vFirstDayOfThisMonth, "B")); //當月最後一天
            //宣告用來存放 SQL 查詢語法的變數
            string vSelectSQLStr = "";
            //取得報表名稱
            lbReportTitle.Text = eReportList.SelectedItem.ToString();
            //取得 SQL 主機的連結字串
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            //================================================================================
            string vRecordDateStr = eStopDate_Year.Text.Trim() + "年" + eStopDate_Month.Text.Trim() + "月";
            string vRecordNote = "查詢資料：固定資產報表" + Environment.NewLine +
                                 "AssetReportList.aspx" + Environment.NewLine +
                                 "報表名稱：" + lbReportTitle.Text + Environment.NewLine +
                                 "截止年月：" + vRecordDateStr;
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            //================================================================================
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp;
                SqlDataAdapter daTemp;
                DataSet dsTemp = new DataSet();
                switch (eReportList.SelectedValue)
                {
                    case "AssetP001":
                        vSelectSQLStr = "select SerialNo, [Name], Unit, AssetNo, [GetDate], [Price], " + Environment.NewLine +
                                        "       AMT1, AMT2, (Price + AMT1 + AMT2) as AMT3, Durable_Years, AMT4, " + Environment.NewLine +
                                        "       AMT5, AMT6, (Price + AMT1 + AMT2 - AMT6) as AMT7 " +
                                        "  from (" + Environment.NewLine +
                                        "        select SerialNo, [Name], " + Environment.NewLine +
                                        "               (select DepNo from AssetB where AssetB.SerialNo = a.SerialNo) unit, " + Environment.NewLine +
                                        "               AssetNo, [GetDate], " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else Price end) Price, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else (select isnull(sum(FixFees), 0) FixFees from AssetX where AssetNo = a.AssetNo and AssetXClass = '3' and BuDate <= '" + vEndOfLastYear.ToString("yyyy/MM/dd") + "') end) AMT1, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else (select isnull(sum(FixFees), 0) FixFees from AssetX where AssetNo = a.AssetNo and AssetXClass = '3' and BuDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "' and BuDate <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') end) AMT2, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else Durable_Years end) Durable_Years, " + Environment.NewLine +
                                        "               (select isnull(sum(Month_Dep), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate >= '" + vFirstDayOfThisMonth.ToString("yyyy/MM/dd") + "' and BuDate <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') AMT4, " + Environment.NewLine +
                                        "               (select isnull(sum(Month_Dep), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "' and BuDate <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') AMT5, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else (select isnull(sum(Month_Dep), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') end) AMT6 " + Environment.NewLine +
                                        "          from Asset a " + Environment.NewLine +
                                        "         where a.Asset_Category = '1' and a.GetDate <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                                        "           and ((a.SellDate is null and a.JunkDate is null) or (a.SellDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "') or (a.JunkDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "')) " + Environment.NewLine +
                                        "        ) z " + Environment.NewLine +
                                        " order by z.[Name], z.AssetNo ";
                        cmdTemp = new SqlCommand(vSelectSQLStr, connTemp);
                        connTemp.Open();
                        daTemp = new SqlDataAdapter(cmdTemp);
                        dsTemp.Clear();
                        daTemp.Fill(dsTemp, eReportList.SelectedValue);
                        gridAssetP001.DataSource = dsTemp;
                        gridAssetP001.DataMember = eReportList.SelectedValue;
                        gridAssetP001.DataBind();
                        plAssetP001.Visible = true;
                        plShowData.Visible = true;
                        break;
                    case "AssetP002":
                        vSelectSQLStr = "select SerialNo, AssetNo, [Name], Unit, [GetDate], Price, RevalValue, " + Environment.NewLine +
                                        "       (Price + isnull(RevalValue, 0)) as AMT1, Durable_Years, SalvageValue, AMT2, AMT3, " + Environment.NewLine +
                                        "       AMT4, ((Price + isnull(RevalValue, 0)) - AMT4) as AMT5 " + Environment.NewLine +
                                        "  from (" + Environment.NewLine +
                                        "        select SerialNo, [Name], " + Environment.NewLine +
                                        "               (select DepNo from AssetB where AssetB.SerialNo = a.SerialNo) unit, " + Environment.NewLine +
                                        "               AssetNo, [GetDate], " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else Price end) Price, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else (select isnull(sum(FixFees), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '3') end) RevalValue, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else Durable_Years end) Durable_Years, " + Environment.NewLine +
                                        "               SalvageValue, " + Environment.NewLine +
                                        "               (select isnull(sum(Month_Dep), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate between '" + vFirstDayOfThisMonth.ToString("yyyy/MM/dd") + "' and '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') AMT2, " + Environment.NewLine +
                                        "               (select isnull(sum(Month_Dep), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate between '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "' and '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') AMT3, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else (select isnull(sum(Month_Dep), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') end) AMT4 " + Environment.NewLine +
                                        "          from Asset a " + Environment.NewLine +
                                        "         where Asset_Category = '2' and a.[GetDate] <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                                        "           and ((a.SellDate is null and a.JunkDate is null) or (a.SellDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "') or (a.JunkDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "')) " + Environment.NewLine +
                                        "        ) z " + Environment.NewLine +
                                        " order by z.AssetNo, z.[Name], z.Unit ";
                        cmdTemp = new SqlCommand(vSelectSQLStr, connTemp);
                        connTemp.Open();
                        daTemp = new SqlDataAdapter(cmdTemp);
                        dsTemp.Clear();
                        daTemp.Fill(dsTemp, eReportList.SelectedValue);
                        gridAssetP002.DataSource = dsTemp;
                        gridAssetP002.DataMember = eReportList.SelectedValue;
                        gridAssetP002.DataBind();
                        plAssetP002.Visible = true;
                        plShowData.Visible = true;
                        break;
                    case "AssetP003":
                        vSelectSQLStr = "select SerialNo, AssetNo, [Name], Unit, [GetDate], " + Environment.NewLine +
                                        "       Price, RevalValue, (Price + isnull(RevalValue, 0)) as AMT1, Durable_Years, SalvageValue, " + Environment.NewLine +
                                        "       AMT2, AMT3, AMT4, ((Price + isnull(RevalValue, 0)) - AMT4) as AMT5 " + Environment.NewLine +
                                        "  from (" + Environment.NewLine +
                                        "        select SerialNo, [Name], " + Environment.NewLine +
                                        "               (select DepNo from AssetB where AssetB.SerialNo = a.SerialNo) unit, " + Environment.NewLine +
                                        "               AssetNo, [GetDate], " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else Price end) Price, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else (select isnull(sum(FixFees), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '3') end) RevalValue, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else Durable_Years end) Durable_Years, " + Environment.NewLine +
                                        "               SalvageValue, " + Environment.NewLine +
                                        "               (select isnull(sum(Month_Dep), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate between '" + vFirstDayOfThisMonth.ToString("yyyy/MM/dd") + "' and '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') AMT2, " + Environment.NewLine +
                                        "               (select isnull(sum(Month_Dep), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate between '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "' and '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') AMT3, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else (select isnull(sum(Month_Dep), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') end) AMT4 " + Environment.NewLine +
                                        "          from Asset a " + Environment.NewLine +
                                        "         where Asset_Category = '3' and a.[GetDate] <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                                        "           and ((a.SellDate is null and a.JunkDate is null) or (a.SellDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "') or (a.JunkDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "')) " + Environment.NewLine +
                                        "        ) z " + Environment.NewLine +
                                        " order by z.AssetNo, z.[Name], z.Unit ";
                        cmdTemp = new SqlCommand(vSelectSQLStr, connTemp);
                        connTemp.Open();
                        daTemp = new SqlDataAdapter(cmdTemp);
                        dsTemp.Clear();
                        daTemp.Fill(dsTemp, eReportList.SelectedValue);
                        gridAssetP003.DataSource = dsTemp;
                        gridAssetP003.DataMember = eReportList.SelectedValue;
                        gridAssetP003.DataBind();
                        plAssetP003.Visible = true;
                        plShowData.Visible = true;
                        break;
                    case "AssetP004":
                        vSelectSQLStr = "select SerialNo, AssetNo, [Name], Unit, [GetDate], Price, Years, Month_Dep, " + Environment.NewLine +
                                        "       AMT1, Addup_Dep, (Price - Addup_Dep) as AMT2, Subject2, SubjectName" + Environment.NewLine +
                                        "  from (" + Environment.NewLine +
                                        "        select SerialNo, [Name], " + Environment.NewLine +
                                        "               (select DepNo from AssetB where AssetB.SerialNo = a.SerialNo) unit, " + Environment.NewLine +
                                        "               AssetNo, [GetDate], " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else Price end) Price, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else Durable_Years end) Years, " + Environment.NewLine +
                                        "               (select isnull(sum(Month_Dep), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate between '" + vFirstDayOfThisMonth.ToString("yyyy/MM/dd") + "' and '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') Month_Dep, " + Environment.NewLine +
                                        "               (select isnull(sum(Month_Dep), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate between '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "' and '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') AMT1, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else Addup_Dep end) Addup_Dep, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else BookValue end) BookValue, " + Environment.NewLine +
                                        "               (Subject2 + isnull((select top 1 DepNo from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "' order by BuDate DESC), '')) Subject2, " + Environment.NewLine +
                                        "               (select [Name] from AC_Subject where [Subject] = a.Subject2 + isnull((select top 1 DepNo from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "' order by BuDate DESC), '')) SubjectName " + Environment.NewLine +
                                        "          from Asset a " + Environment.NewLine +
                                        "         where Asset_Category = '4' and a.[GetDate] <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                                        "           and ((a.SellDate is null and a.JunkDate is null) or (a.SellDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "') or (a.JunkDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "')) " + Environment.NewLine +
                                        "        ) z " + Environment.NewLine +
                                        " order by z.Subject2, z.AssetNo, z.[Name] ";
                        cmdTemp = new SqlCommand(vSelectSQLStr, connTemp);
                        connTemp.Open();
                        daTemp = new SqlDataAdapter(cmdTemp);
                        dsTemp.Clear();
                        daTemp.Fill(dsTemp, eReportList.SelectedValue);
                        gridAssetP004.DataSource = dsTemp;
                        gridAssetP004.DataMember = eReportList.SelectedValue;
                        gridAssetP004.DataBind();
                        plAssetP004.Visible = true;
                        plShowData.Visible = true;
                        break;
                    case "AssetP005":
                        vSelectSQLStr = "select SerialNo, AssetNo, Unit, ([Name] + ' (' + cast([Area] as varchar) + ') ') as [Name], " + Environment.NewLine +
                                        "       [Location], [GetDate], Price, AMT1, AMT2, (Price + AMT1 + AMT2) as AMT3, " + Environment.NewLine +
                                        "       Durable_Years, AMT4, AMT5, AMT6, ((Price + AMT1 + AMT2) - AMT6) as AMT7 " + Environment.NewLine +
                                        "  from (" + Environment.NewLine +
                                        "        select SerialNo, AssetNo, " + Environment.NewLine +
                                        "               (select DepNo from AssetB where AssetB.SerialNo = a.SerialNo) Unit, " + Environment.NewLine +
                                        "               [Name], [Area], [Location], [GetDate], " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else Price end) Price, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else (select isnull(sum(FixFees), 0) FixFees from AssetX where AssetNo = a.AssetNo and AssetXClass = '3' and BuDate <= '" + vEndOfLastYear.ToString("yyyy/MM/dd") + "') end) AMT1, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else (select isnull(sum(FixFees), 0) FixFees from AssetX where AssetNo = a.AssetNo and AssetXClass = '3' and BuDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "' and BuDate <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') end) AMT2, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else Durable_Years end) Durable_Years, " + Environment.NewLine +
                                        "               (select isnull(sum(Month_Dep), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate >= '" + vFirstDayOfThisMonth.ToString("yyyy/MM/dd") + "' and BuDate <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') AMT4, " + Environment.NewLine +
                                        "               (select isnull(sum(Month_Dep), 0) from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "' and BuDate <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') AMT5, " + Environment.NewLine +
                                        "               (case when ((SellDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "') or (JunkDate < '" + vFirstDayOfNextMonth.ToString("yyyy/MM/dd") + "')) then 0 else (select isnull(sum(Month_Dep), 0) Month_Dep from AssetX where AssetNo = a.AssetNo and AssetXClass = '2' and BuDate <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "') end) AMT6 " + Environment.NewLine +
                                        "          from Asset a " + Environment.NewLine +
                                        "         where Asset_Category = '5' and a.[GetDate] <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                                        "           and ((a.SellDate is null and a.JunkDate is null) or (a.SellDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "') or (a.JunkDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "')) " + Environment.NewLine +
                                        "        ) z " + Environment.NewLine +
                                        " order by z.AssetNo, z.[Name], z.Unit ";
                        cmdTemp = new SqlCommand(vSelectSQLStr, connTemp);
                        connTemp.Open();
                        daTemp = new SqlDataAdapter(cmdTemp);
                        dsTemp.Clear();
                        daTemp.Fill(dsTemp, eReportList.SelectedValue);
                        gridAssetP005.DataSource = dsTemp;
                        gridAssetP005.DataMember = eReportList.SelectedValue;
                        gridAssetP005.DataBind();
                        plAssetP005.Visible = true;
                        plShowData.Visible = true;
                        break;
                    case "AssetP006":
                        vSelectSQLStr = "select SerialNo, AssetNo, SellDate, JunkDate, [Location], [Area], [GetDate], Price, " + Environment.NewLine +
                                        "       L_N_Amount, L_IncTax_Prepare, (L_N_Amount - L_IncTax_Prepare) as AMT1 " + Environment.NewLine +
                                        "  from Asset a " + Environment.NewLine +
                                        " where Asset_Category = '6' and State in ('0', '3') and a.[GetDate] <= '" + vLastDayOfThisMonth.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                                        "   and ((a.SellDate is null and a.JunkDate is null) or (a.SellDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "') or (a.JunkDate >= '" + vFirstDayOfThisYear.ToString("yyyy/MM/dd") + "')) " + Environment.NewLine +
                                        " order by a.[Name], a.AssetNo ";
                        cmdTemp = new SqlCommand(vSelectSQLStr, connTemp);
                        connTemp.Open();
                        daTemp = new SqlDataAdapter(cmdTemp);
                        dsTemp.Clear();
                        daTemp.Fill(dsTemp, eReportList.SelectedValue);
                        gridAssetP006.DataSource = dsTemp;
                        gridAssetP006.DataMember = eReportList.SelectedValue;
                        gridAssetP006.DataBind();
                        plAssetP006.Visible = true;
                        plShowData.Visible = true;
                        break;
                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbLayout_Click(object sender, EventArgs e)
        {
            bool vLayoutStatus = false;
            string vSavePath = eReportList.SelectedValue;
            //================================================================================
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vRecordDateStr = eStopDate_Year.Text.Trim() + "年" + eStopDate_Month.Text.Trim() + "月";
            string vRecordNote = "匯出檔案：固定資產報表" + Environment.NewLine +
                                 "AssetReportList.aspx" + Environment.NewLine +
                                 "報表名稱：" + eReportList.SelectedItem.ToString() + Environment.NewLine +
                                 "截止年月：" + vRecordDateStr;
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            //================================================================================
            switch (vSavePath)
            {
                case "AssetP001":
                    vLayoutStatus = SaveToExcel_P001(lbReportTitle.Text.Trim(), vSavePath);
                    break;
                case "AssetP002":
                    vLayoutStatus = SaveToExcel_P002(lbReportTitle.Text.Trim(), vSavePath);
                    break;
                case "AssetP003":
                    vLayoutStatus = SaveToExcel_P003(lbReportTitle.Text.Trim(), vSavePath);
                    break;
                case "AssetP004":
                    vLayoutStatus = SaveToExcel_P004(lbReportTitle.Text.Trim(), vSavePath);
                    break;
                case "AssetP005":
                    vLayoutStatus = SaveToExcel_P005(lbReportTitle.Text.Trim(), vSavePath);
                    break;
                case "AssetP006":
                    vLayoutStatus = SaveToExcel_P006(lbReportTitle.Text.Trim(), vSavePath);
                    break;
            }
            if (!vLayoutStatus)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('匯出資料失敗，請洽電腦課人員！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbClearLayout_Click(object sender, EventArgs e)
        {
            plAssetP001.Visible = false;
            plAssetP002.Visible = false;
            plAssetP003.Visible = false;
            plAssetP004.Visible = false;
            plAssetP005.Visible = false;
            plAssetP006.Visible = false;
            plShowData.Visible = false;
        }

        private bool SaveToExcel_P001(string fRPTName, string fSavePath)
        {
            bool vResultStatus = false;
            int vColumnCount = gridAssetP001.Columns.Count;
            string vPrintDate = ((DateTime.Now.Year) - 1911).ToString("D3") + "/" + DateTime.Now.ToString("MM/dd");
            string vGetDate = "";
            DateTime vTempDate;
            //Excel 檔案
            HSSFWorkbook wbExcel = new HSSFWorkbook();
            //Excel 工作表
            HSSFSheet wsExcel;
            //設定標題欄位的格式
            HSSFCellStyle csTitle = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中
                                                            //設定字體格式
            HSSFFont fontTitle = (HSSFFont)wbExcel.CreateFont();
            //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
            fontTitle.IsBold = true;
            fontTitle.FontHeightInPoints = 20; //字體大小
            csTitle.SetFont(fontTitle);
            //換頁用的欄位值
            string vGroupText = "";
            int vGroupIndex = 0;
            //總計用的公式
            string vFormula_CarCount = "sum(";
            string vFormula_AMT3 = "sum(";
            string vFormula_AMT4 = "sum(";
            string vFormula_AMT5 = "sum(";
            string vFormula_AMT6 = "sum(";
            string vFormula_AMT7 = "sum(";

            foreach (DataControlField vCol in gridAssetP001.Columns)
            {
                if (vCol.HeaderText.Trim() == "廠牌")
                {
                    vGroupIndex = gridAssetP001.Columns.IndexOf(vCol);
                }
            }
            int vLinesNo = 0;

            //開始建立資料
            for (int i = 0; i < gridAssetP001.Rows.Count; i++)
            {
                //換頁鍵值有變的時候
                if (gridAssetP001.Rows[i].Cells[vGroupIndex].Text.Trim() != vGroupText)
                {
                    if (vLinesNo != 0) //行數不是 0 表示有上一頁
                    {
                        wsExcel = (HSSFSheet)wbExcel.GetSheet(vGroupText);
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("廠牌合計：");
                        wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("車輛數：");
                        wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellFormula("count(G5:G" + vLinesNo.ToString() + ")");
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue(" 輛");
                        wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellFormula("sum(I5:I" + vLinesNo.ToString() + ")");
                        wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellFormula("sum(K5:K" + vLinesNo.ToString() + ")");
                        wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellFormula("sum(L5:L" + vLinesNo.ToString() + ")");
                        wsExcel.GetRow(vLinesNo).CreateCell(12).SetCellFormula("sum(M5:M" + vLinesNo.ToString() + ")");
                        wsExcel.GetRow(vLinesNo).CreateCell(13).SetCellFormula("sum(N5:N" + vLinesNo.ToString() + ")");
                        //組合總計公式
                        vFormula_CarCount = vFormula_CarCount + vGroupText + "!" + "G" + (vLinesNo + 1).ToString() + ",";
                        vFormula_AMT3 = vFormula_AMT3 + vGroupText + "!" + "I" + (vLinesNo + 1).ToString() + ",";
                        vFormula_AMT4 = vFormula_AMT4 + vGroupText + "!" + "K" + (vLinesNo + 1).ToString() + ",";
                        vFormula_AMT5 = vFormula_AMT5 + vGroupText + "!" + "L" + (vLinesNo + 1).ToString() + ",";
                        vFormula_AMT6 = vFormula_AMT6 + vGroupText + "!" + "M" + (vLinesNo + 1).ToString() + ",";
                        vFormula_AMT7 = vFormula_AMT7 + vGroupText + "!" + "N" + (vLinesNo + 1).ToString() + ",";
                    }

                    vLinesNo = 0;
                    vGroupText = gridAssetP001.Rows[i].Cells[vGroupIndex].Text.Trim();
                    //Excel 工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vGroupText);

                    //建立報表表頭區
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("桃園汽車客運股份有限公司");
                    //wsExcel.AutoSizeColumn(0);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, vColumnCount - 1));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("");
                    //wsExcel.AutoSizeColumn(0);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 1));
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue(fRPTName);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, vColumnCount - 3));
                    wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csTitle;
                    wsExcel.GetRow(vLinesNo).CreateCell(vColumnCount - 2).SetCellValue("印表日期：" + vPrintDate);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, vColumnCount - 2, vColumnCount - 1));
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue((int.Parse(eStopDate_Year.Text.Trim()) - 1911).ToString() + "年");
                    //wsExcel.AutoSizeColumn(0);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, vColumnCount - 1));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                    vLinesNo++;

                    //設定標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int j = 0; j < gridAssetP001.Columns.Count; j++)
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(j).SetCellValue(gridAssetP001.Columns[j].HeaderText);
                    }
                    vLinesNo++;
                }
                //寫入內容
                wsExcel = (HSSFSheet)wbExcel.GetSheet(vGroupText);
                wsExcel.CreateRow(vLinesNo);
                for (int k = 0; k < gridAssetP001.Columns.Count; k++)
                {
                    if ((k >= 5) && (k <= 13)) //這幾個欄位是數值欄位
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(k);
                        wsExcel.GetRow(vLinesNo).GetCell(k).SetCellType(CellType.Numeric);
                        wsExcel.GetRow(vLinesNo).GetCell(k).SetCellValue(double.Parse(gridAssetP001.Rows[i].Cells[k].Text.Trim()));
                    }
                    else
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(k);
                        wsExcel.GetRow(vLinesNo).GetCell(k).SetCellType(CellType.String);
                        if (gridAssetP001.Columns[k].HeaderText == "取得日期")
                        {
                            vTempDate = DateTime.Parse(gridAssetP001.Rows[i].Cells[k].Text.Trim());
                            vGetDate = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                            wsExcel.GetRow(vLinesNo).GetCell(k).SetCellValue(vGetDate);
                        }
                        else
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(k).SetCellValue(gridAssetP001.Rows[i].Cells[k].Text.Trim());
                        }
                    }

                }
                vLinesNo++;
            }
            wsExcel = (HSSFSheet)wbExcel.GetSheet(vGroupText);
            wsExcel.CreateRow(vLinesNo);
            wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("廠牌合計：");
            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("車輛數：");
            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellFormula("count(G5:G" + vLinesNo.ToString() + ")");
            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue(" 輛");
            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellFormula("sum(I5:I" + vLinesNo.ToString() + ")");
            wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellFormula("sum(K5:K" + vLinesNo.ToString() + ")");
            wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellFormula("sum(L5:L" + vLinesNo.ToString() + ")");
            wsExcel.GetRow(vLinesNo).CreateCell(12).SetCellFormula("sum(M5:M" + vLinesNo.ToString() + ")");
            wsExcel.GetRow(vLinesNo).CreateCell(13).SetCellFormula("sum(N5:N" + vLinesNo.ToString() + ")");
            //組合總計公式
            vFormula_CarCount = vFormula_CarCount + vGroupText + "!" + "G" + (vLinesNo + 1).ToString() + ")";
            vFormula_AMT3 = vFormula_AMT3 + vGroupText + "!" + "I" + (vLinesNo + 1).ToString() + ")";
            vFormula_AMT4 = vFormula_AMT4 + vGroupText + "!" + "K" + (vLinesNo + 1).ToString() + ")";
            vFormula_AMT5 = vFormula_AMT5 + vGroupText + "!" + "L" + (vLinesNo + 1).ToString() + ")";
            vFormula_AMT6 = vFormula_AMT6 + vGroupText + "!" + "M" + (vLinesNo + 1).ToString() + ")";
            vFormula_AMT7 = vFormula_AMT7 + vGroupText + "!" + "N" + (vLinesNo + 1).ToString() + ")";
            vLinesNo++;
            //寫入總計
            wsExcel.CreateRow(vLinesNo);
            wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("總　　計：");
            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("車輛數：");
            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellFormula(vFormula_CarCount);
            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue(" 輛");
            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellFormula(vFormula_AMT3);
            wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellFormula(vFormula_AMT4);
            wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellFormula(vFormula_AMT5);
            wsExcel.GetRow(vLinesNo).CreateCell(12).SetCellFormula(vFormula_AMT6);
            wsExcel.GetRow(vLinesNo).CreateCell(13).SetCellFormula(vFormula_AMT7);

            try
            {
                /*
                MemoryStream msTarget = new MemoryStream();
                wbExcel.Write(msTarget);
                // 設定強制下載標頭
                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fSavePath + ".xls"));
                // 輸出檔案
                Response.BinaryWrite(msTarget.ToArray());
                msTarget.Close();

                Response.End();
                vResultStatus = true; */
                var msTarget = new NPOIMemoryStream();
                msTarget.AllowClose = false;
                wbExcel.Write(msTarget);
                msTarget.Flush();
                msTarget.Seek(0, SeekOrigin.Begin);
                msTarget.AllowClose = true;

                if (msTarget.Length > 0)
                {
                    //先判斷是不是IE
                    HttpContext.Current.Response.ContentType = "application/octet-stream";
                    HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                    string TourVerision = brObject.Type;
                    if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                    {
                        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + fSavePath + ".xls", System.Text.Encoding.UTF8));
                    }
                    else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                    {
                        // 設定強制下載標頭
                        Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fSavePath + ".xls"));
                    }
                    // 輸出檔案
                    Response.BinaryWrite(msTarget.ToArray());
                    msTarget.Close();

                    Response.End();
                    vResultStatus = true;
                }
            }
            catch (Exception eMessage)
            {
                vResultStatus = false;
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
            return vResultStatus;
        }

        private bool SaveToExcel_P002(string fRPTName, string fSavePath)
        {
            bool vResultStatus = false;
            int vColumnCount = gridAssetP002.Columns.Count;
            string vPrintDate = ((DateTime.Now.Year) - 1911).ToString("D3") + "/" + DateTime.Now.ToString("MM/dd");
            string vGetDate = "";
            DateTime vTempDate;
            //Excel 檔案
            HSSFWorkbook wbExcel = new HSSFWorkbook();
            //Excel 工作表
            HSSFSheet wsExcel = (HSSFSheet)wbExcel.CreateSheet(fRPTName);
            //設定標題欄位的格式
            HSSFCellStyle csTitle = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中
                                                            //設定字體格式
            HSSFFont fontTitle = (HSSFFont)wbExcel.CreateFont();
            //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
            fontTitle.IsBold = true;
            fontTitle.FontHeightInPoints = 20; //字體大小
            csTitle.SetFont(fontTitle);

            //建立報表表頭區
            wsExcel.CreateRow(0);
            wsExcel.GetRow(0).CreateCell(0).SetCellValue("桃園汽車客運股份有限公司");
            //wsExcel.AutoSizeColumn(0);
            wsExcel.AddMergedRegion(new CellRangeAddress(0, 0, 0, vColumnCount - 1));
            wsExcel.GetRow(0).GetCell(0).CellStyle = csTitle;
            wsExcel.CreateRow(1);
            wsExcel.GetRow(1).CreateCell(0).SetCellValue("");
            //wsExcel.AutoSizeColumn(0);
            wsExcel.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1));
            wsExcel.GetRow(1).CreateCell(2).SetCellValue(fRPTName);
            wsExcel.AddMergedRegion(new CellRangeAddress(1, 1, 2, vColumnCount - 3));
            wsExcel.GetRow(1).GetCell(2).CellStyle = csTitle;
            wsExcel.GetRow(1).CreateCell(vColumnCount - 2).SetCellValue("印表日期：" + vPrintDate);
            wsExcel.AddMergedRegion(new CellRangeAddress(1, 1, vColumnCount - 2, vColumnCount - 1));
            wsExcel.CreateRow(2);
            wsExcel.GetRow(2).CreateCell(0).SetCellValue((int.Parse(eStopDate_Year.Text.Trim()) - 1911).ToString() + "年");
            //wsExcel.AutoSizeColumn(0);
            wsExcel.AddMergedRegion(new CellRangeAddress(2, 2, 0, vColumnCount - 1));
            wsExcel.GetRow(2).GetCell(0).CellStyle = csTitle;

            //建立報表內容
            //第一行是標題列
            wsExcel.CreateRow(3);
            for (int i = 0; i < gridAssetP002.Columns.Count; i++)
            {
                wsExcel.GetRow(3).CreateCell(i).SetCellValue(gridAssetP002.Columns[i].HeaderText);
            }
            //寫入內容
            for (int j = 0; j < gridAssetP002.Rows.Count; j++)
            {
                wsExcel.CreateRow(j + 4);
                for (int k = 0; k < gridAssetP002.Columns.Count; k++)
                {
                    if ((k >= 5) && (k <= 13)) //這幾個欄位是數值欄位
                    {
                        wsExcel.GetRow(j + 4).CreateCell(k);
                        wsExcel.GetRow(j + 4).GetCell(k).SetCellType(CellType.Numeric);
                        wsExcel.GetRow(j + 4).GetCell(k).SetCellValue(double.Parse(gridAssetP002.Rows[j].Cells[k].Text.Trim()));
                    }
                    else
                    {
                        wsExcel.GetRow(j + 4).CreateCell(k);
                        wsExcel.GetRow(j + 4).GetCell(k).SetCellType(CellType.String);
                        if (gridAssetP002.Columns[k].HeaderText == "取得日期")
                        {
                            vTempDate = DateTime.Parse(gridAssetP002.Rows[j].Cells[k].Text.Trim());
                            vGetDate = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                            wsExcel.GetRow(j + 4).GetCell(k).SetCellValue(vGetDate);
                        }
                        else
                        {
                            wsExcel.GetRow(j + 4).GetCell(k).SetCellValue(gridAssetP002.Rows[j].Cells[k].Text.Trim());
                        }
                    }
                }
            }
            //寫入小計欄位
            int TotalRows = gridAssetP002.Rows.Count + 4;
            wsExcel.CreateRow(TotalRows);
            wsExcel.GetRow(TotalRows).CreateCell(6).SetCellValue("合計：");
            wsExcel.GetRow(TotalRows).CreateCell(7).SetCellFormula("sum(H5:H" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(10).SetCellFormula("sum(K5:K" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(11).SetCellFormula("sum(L5:L" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(12).SetCellFormula("sum(M5:M" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(13).SetCellFormula("sum(N5:N" + TotalRows.ToString() + ")");

            try
            {
                /*
                MemoryStream msTarget = new MemoryStream();
                wbExcel.Write(msTarget);
                // 設定強制下載標頭
                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fSavePath + ".xls"));
                // 輸出檔案
                Response.BinaryWrite(msTarget.ToArray());
                msTarget.Close();

                Response.End();
                vResultStatus = true; //*/
                var msTarget = new NPOIMemoryStream();
                msTarget.AllowClose = false;
                wbExcel.Write(msTarget);
                msTarget.Flush();
                msTarget.Seek(0, SeekOrigin.Begin);
                msTarget.AllowClose = true;

                if (msTarget.Length > 0)
                {
                    //先判斷是不是IE
                    HttpContext.Current.Response.ContentType = "application/octet-stream";
                    HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                    string TourVerision = brObject.Type;
                    if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                    {
                        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + fSavePath + ".xls", System.Text.Encoding.UTF8));
                    }
                    else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                    {
                        // 設定強制下載標頭
                        Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fSavePath + ".xls"));
                    }
                    // 輸出檔案
                    Response.BinaryWrite(msTarget.ToArray());
                    msTarget.Close();

                    Response.End();
                    vResultStatus = true;
                }
            }
            catch (Exception eMessage)
            {
                vResultStatus = false;
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
            return vResultStatus;
        }

        private bool SaveToExcel_P003(string fRPTName, string fSavePath)
        {
            bool vResultStatus = false;
            int vColumnCount = gridAssetP003.Columns.Count;
            string vPrintDate = ((DateTime.Now.Year) - 1911).ToString("D3") + "/" + DateTime.Now.ToString("MM/dd");
            string vGetDate = "";
            DateTime vTempDate;
            //Excel 檔案
            HSSFWorkbook wbExcel = new HSSFWorkbook();
            //Excel 工作表
            HSSFSheet wsExcel = (HSSFSheet)wbExcel.CreateSheet(fRPTName);
            //設定標題欄位的格式
            HSSFCellStyle csTitle = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中
                                                            //設定字體格式
            HSSFFont fontTitle = (HSSFFont)wbExcel.CreateFont();
            //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
            fontTitle.IsBold = true;
            fontTitle.FontHeightInPoints = 20; //字體大小
            csTitle.SetFont(fontTitle);

            //建立報表表頭區
            wsExcel.CreateRow(0);
            wsExcel.GetRow(0).CreateCell(0).SetCellValue("桃園汽車客運股份有限公司");
            //wsExcel.AutoSizeColumn(0);
            wsExcel.AddMergedRegion(new CellRangeAddress(0, 0, 0, vColumnCount - 1));
            wsExcel.GetRow(0).GetCell(0).CellStyle = csTitle;
            wsExcel.CreateRow(1);
            wsExcel.GetRow(1).CreateCell(0).SetCellValue("");
            //wsExcel.AutoSizeColumn(0);
            wsExcel.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1));
            wsExcel.GetRow(1).CreateCell(2).SetCellValue(fRPTName);
            wsExcel.AddMergedRegion(new CellRangeAddress(1, 1, 2, vColumnCount - 3));
            wsExcel.GetRow(1).GetCell(2).CellStyle = csTitle;
            wsExcel.GetRow(1).CreateCell(vColumnCount - 2).SetCellValue("印表日期：" + vPrintDate);
            wsExcel.AddMergedRegion(new CellRangeAddress(1, 1, vColumnCount - 2, vColumnCount - 1));
            wsExcel.CreateRow(2);
            wsExcel.GetRow(2).CreateCell(0).SetCellValue((int.Parse(eStopDate_Year.Text.Trim()) - 1911).ToString() + "年");
            //wsExcel.AutoSizeColumn(0);
            wsExcel.AddMergedRegion(new CellRangeAddress(2, 2, 0, vColumnCount - 1));
            wsExcel.GetRow(2).GetCell(0).CellStyle = csTitle;

            //建立報表內容
            //第一行是標題列
            wsExcel.CreateRow(3);
            for (int i = 0; i < gridAssetP003.Columns.Count; i++)
            {
                wsExcel.GetRow(3).CreateCell(i).SetCellValue(gridAssetP003.Columns[i].HeaderText);
            }
            //寫入內容
            for (int j = 0; j < gridAssetP003.Rows.Count; j++)
            {
                wsExcel.CreateRow(j + 4);
                for (int k = 0; k < gridAssetP003.Columns.Count; k++)
                {
                    if ((k >= 5) && (k <= 13)) //這幾個欄位是數值欄位
                    {
                        wsExcel.GetRow(j + 4).CreateCell(k);
                        wsExcel.GetRow(j + 4).GetCell(k).SetCellType(CellType.Numeric);
                        wsExcel.GetRow(j + 4).GetCell(k).SetCellValue(double.Parse(gridAssetP003.Rows[j].Cells[k].Text.Trim()));
                    }
                    else
                    {
                        wsExcel.GetRow(j + 4).CreateCell(k);
                        wsExcel.GetRow(j + 4).GetCell(k).SetCellType(CellType.String);
                        if (gridAssetP003.Columns[k].HeaderText == "取得日期")
                        {
                            vTempDate = DateTime.Parse(gridAssetP003.Rows[j].Cells[k].Text.Trim());
                            vGetDate = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                            wsExcel.GetRow(j + 4).GetCell(k).SetCellValue(vGetDate);
                        }
                        else
                        {
                            wsExcel.GetRow(j + 4).GetCell(k).SetCellValue(gridAssetP003.Rows[j].Cells[k].Text.Trim());
                        }
                    }
                }
            }
            //寫入小計欄位
            int TotalRows = gridAssetP003.Rows.Count + 4;
            wsExcel.CreateRow(TotalRows);
            wsExcel.GetRow(TotalRows).CreateCell(6).SetCellValue("合計：");
            wsExcel.GetRow(TotalRows).CreateCell(7).SetCellFormula("sum(H5:H" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(10).SetCellFormula("sum(K5:K" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(11).SetCellFormula("sum(L5:L" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(12).SetCellFormula("sum(M5:M" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(13).SetCellFormula("sum(N5:N" + TotalRows.ToString() + ")");

            try
            {
                /*
                MemoryStream msTarget = new MemoryStream();
                wbExcel.Write(msTarget);
                // 設定強制下載標頭
                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fSavePath + ".xls"));
                // 輸出檔案
                Response.BinaryWrite(msTarget.ToArray());
                msTarget.Close();

                Response.End();
                vResultStatus = true;//*/
                var msTarget = new NPOIMemoryStream();
                msTarget.AllowClose = false;
                wbExcel.Write(msTarget);
                msTarget.Flush();
                msTarget.Seek(0, SeekOrigin.Begin);
                msTarget.AllowClose = true;

                if (msTarget.Length > 0)
                {
                    //先判斷是不是IE
                    HttpContext.Current.Response.ContentType = "application/octet-stream";
                    HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                    string TourVerision = brObject.Type;
                    if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                    {
                        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + fSavePath + ".xls", System.Text.Encoding.UTF8));
                    }
                    else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                    {
                        // 設定強制下載標頭
                        Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fSavePath + ".xls"));
                    }
                    // 輸出檔案
                    Response.BinaryWrite(msTarget.ToArray());
                    msTarget.Close();

                    Response.End();
                    vResultStatus = true;
                }
            }
            catch (Exception eMessage)
            {
                vResultStatus = false;
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
            return vResultStatus;
        }

        private bool SaveToExcel_P004(string fRPTName, string fSavePath)
        {
            bool vResultStatus = false;
            int vColumnCount = gridAssetP004.Columns.Count;
            string vPrintDate = ((DateTime.Now.Year) - 1911).ToString("D3") + "/" + DateTime.Now.ToString("MM/dd");
            string vGetDate = "";
            DateTime vTempDate;
            //Excel 檔案
            HSSFWorkbook wbExcel = new HSSFWorkbook();
            //Excel 工作表
            HSSFSheet wsExcel;
            //設定標題欄位的格式
            HSSFCellStyle csTitle = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中
                                                            //設定字體格式
            HSSFFont fontTitle = (HSSFFont)wbExcel.CreateFont();
            //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
            fontTitle.IsBold = true;
            fontTitle.FontHeightInPoints = 20; //字體大小
            csTitle.SetFont(fontTitle);
            //換頁用的欄位值
            string vGroupText = "";
            int vGroupIndex = 0;
            //總計用的公式
            string vFormula_Price = "sum(";
            string vFormula_Month_Dep = "sum(";
            string vFormula_AMT1 = "sum(";
            string vFormula_Addup_Dep = "sum(";
            string vFormula_AMT2 = "sum(";

            foreach (DataControlField vCol in gridAssetP004.Columns)
            {
                if (vCol.HeaderText.Trim() == "科目代號")
                {
                    vGroupIndex = gridAssetP004.Columns.IndexOf(vCol);
                }
            }
            int vLinesNo = 0;

            //開始建立資料
            for (int i = 0; i < gridAssetP004.Rows.Count; i++)
            {
                //換頁鍵值有變的時候
                if (gridAssetP004.Rows[i].Cells[vGroupIndex].Text.Trim() != vGroupText)
                {
                    if (vLinesNo != 0) //行數不是 0 表示有上一頁
                    {
                        wsExcel = (HSSFSheet)wbExcel.GetSheet(vGroupText);
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("科目合計：");
                        wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellFormula("sum(F5:F" + vLinesNo.ToString() + ")");
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellFormula("sum(H5:H" + vLinesNo.ToString() + ")");
                        wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellFormula("sum(I5:I" + vLinesNo.ToString() + ")");
                        wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellFormula("sum(J5:J" + vLinesNo.ToString() + ")");
                        wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellFormula("sum(K5:K" + vLinesNo.ToString() + ")");
                        //組合總計公式
                        vFormula_Price = vFormula_Price + "'" + vGroupText + "'!" + "F" + (vLinesNo + 1).ToString() + ",";
                        vFormula_Month_Dep = vFormula_Month_Dep + "'" + vGroupText + "'!" + "H" + (vLinesNo + 1).ToString() + ",";
                        vFormula_AMT1 = vFormula_AMT1 + "'" + vGroupText + "'!" + "I" + (vLinesNo + 1).ToString() + ",";
                        vFormula_Addup_Dep = vFormula_Addup_Dep + "'" + vGroupText + "'!" + "J" + (vLinesNo + 1).ToString() + ",";
                        vFormula_AMT2 = vFormula_AMT2 + "'" + vGroupText + "'!" + "K" + (vLinesNo + 1).ToString() + ",";
                    }

                    vLinesNo = 0;
                    vGroupText = gridAssetP004.Rows[i].Cells[vGroupIndex].Text.Trim();
                    //Excel 工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vGroupText);

                    //建立報表表頭區
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("桃園汽車客運股份有限公司");
                    //wsExcel.AutoSizeColumn(0);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, vColumnCount - 1));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("");
                    //wsExcel.AutoSizeColumn(0);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 1));
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue(fRPTName);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, vColumnCount - 3));
                    wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csTitle;
                    wsExcel.GetRow(vLinesNo).CreateCell(vColumnCount - 2).SetCellValue("印表日期：" + vPrintDate);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, vColumnCount - 2, vColumnCount - 1));
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue((int.Parse(eStopDate_Year.Text.Trim()) - 1911).ToString() + "年");
                    //wsExcel.AutoSizeColumn(0);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, vColumnCount - 1));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                    vLinesNo++;

                    //設定標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int j = 0; j < gridAssetP004.Columns.Count; j++)
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(j).SetCellValue(gridAssetP004.Columns[j].HeaderText);
                    }
                    vLinesNo++;
                }
                //寫入內容
                wsExcel = (HSSFSheet)wbExcel.GetSheet(vGroupText);
                wsExcel.CreateRow(vLinesNo);
                for (int k = 0; k < gridAssetP004.Columns.Count; k++)
                {
                    if ((k >= 5) && (k <= 10)) //這幾個欄位是數值欄位
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(k);
                        wsExcel.GetRow(vLinesNo).GetCell(k).SetCellType(CellType.Numeric);
                        wsExcel.GetRow(vLinesNo).GetCell(k).SetCellValue(double.Parse(gridAssetP004.Rows[i].Cells[k].Text.Trim()));
                    }
                    else
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(k);
                        wsExcel.GetRow(vLinesNo).GetCell(k).SetCellType(CellType.String);
                        if (gridAssetP004.Columns[k].HeaderText == "取得日期")
                        {
                            vTempDate = DateTime.Parse(gridAssetP004.Rows[i].Cells[k].Text.Trim());
                            vGetDate = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                            wsExcel.GetRow(vLinesNo).GetCell(k).SetCellValue(vGetDate);
                        }
                        else
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(k).SetCellValue(gridAssetP004.Rows[i].Cells[k].Text.Trim());
                        }
                    }
                }
                vLinesNo++;
            }
            wsExcel = (HSSFSheet)wbExcel.GetSheet(vGroupText);
            wsExcel.CreateRow(vLinesNo);
            wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("科目合計：");
            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellFormula("sum(F5:F" + vLinesNo.ToString() + ")");
            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellFormula("sum(H5:H" + vLinesNo.ToString() + ")");
            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellFormula("sum(I5:I" + vLinesNo.ToString() + ")");
            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellFormula("sum(J5:J" + vLinesNo.ToString() + ")");
            wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellFormula("sum(K5:K" + vLinesNo.ToString() + ")");
            //組合總計公式
            vFormula_Price = vFormula_Price + "'" + vGroupText + "'!" + "F" + (vLinesNo + 1).ToString() + ")";
            vFormula_Month_Dep = vFormula_Month_Dep + "'" + vGroupText + "'!" + "H" + (vLinesNo + 1).ToString() + ")";
            vFormula_AMT1 = vFormula_AMT1 + "'" + vGroupText + "'!" + "I" + (vLinesNo + 1).ToString() + ")";
            vFormula_Addup_Dep = vFormula_Addup_Dep + "'" + vGroupText + "'!" + "J" + (vLinesNo + 1).ToString() + ")";
            vFormula_AMT2 = vFormula_AMT2 + "'" + vGroupText + "'!" + "K" + (vLinesNo + 1).ToString() + ")";
            vLinesNo++;
            //寫入總計
            wsExcel.CreateRow(vLinesNo);
            wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("總　　計：");
            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellFormula(vFormula_Price);
            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellFormula(vFormula_Month_Dep);
            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellFormula(vFormula_AMT1);
            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellFormula(vFormula_Addup_Dep);
            wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellFormula(vFormula_AMT2);

            try
            {
                /*
                MemoryStream msTarget = new MemoryStream();
                wbExcel.Write(msTarget);
                // 設定強制下載標頭
                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fSavePath + ".xls"));
                // 輸出檔案
                Response.BinaryWrite(msTarget.ToArray());
                msTarget.Close();

                Response.End();
                vResultStatus = true;//*/
                var msTarget = new NPOIMemoryStream();
                msTarget.AllowClose = false;
                wbExcel.Write(msTarget);
                msTarget.Flush();
                msTarget.Seek(0, SeekOrigin.Begin);
                msTarget.AllowClose = true;

                if (msTarget.Length > 0)
                {
                    //先判斷是不是IE
                    HttpContext.Current.Response.ContentType = "application/octet-stream";
                    HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                    string TourVerision = brObject.Type;
                    if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                    {
                        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + fSavePath + ".xls", System.Text.Encoding.UTF8));
                    }
                    else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                    {
                        // 設定強制下載標頭
                        Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fSavePath + ".xls"));
                    }
                    // 輸出檔案
                    Response.BinaryWrite(msTarget.ToArray());
                    msTarget.Close();

                    Response.End();
                    vResultStatus = true;
                }
            }
            catch (Exception eMessage)
            {
                vResultStatus = false;
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
            return vResultStatus;
        }

        private bool SaveToExcel_P005(string fRPTName, string fSavePath)
        {
            bool vResultStatus = false;
            int vColumnCount = gridAssetP005.Columns.Count;
            string vPrintDate = ((DateTime.Now.Year) - 1911).ToString("D3") + "/" + DateTime.Now.ToString("MM/dd");
            string vGetDate = "";
            DateTime vTempDate;
            //Excel 檔案
            HSSFWorkbook wbExcel = new HSSFWorkbook();
            //Excel 工作表
            HSSFSheet wsExcel = (HSSFSheet)wbExcel.CreateSheet(fRPTName);
            //設定標題欄位的格式
            HSSFCellStyle csTitle = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中
                                                            //設定字體格式
            HSSFFont fontTitle = (HSSFFont)wbExcel.CreateFont();
            //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
            fontTitle.IsBold = true;
            fontTitle.FontHeightInPoints = 20; //字體大小
            csTitle.SetFont(fontTitle);

            //建立報表表頭區
            wsExcel.CreateRow(0);
            wsExcel.GetRow(0).CreateCell(0).SetCellValue("桃園汽車客運股份有限公司");
            //wsExcel.AutoSizeColumn(0);
            wsExcel.AddMergedRegion(new CellRangeAddress(0, 0, 0, vColumnCount - 1));
            wsExcel.GetRow(0).GetCell(0).CellStyle = csTitle;
            wsExcel.CreateRow(1);
            wsExcel.GetRow(1).CreateCell(0).SetCellValue("");
            //wsExcel.AutoSizeColumn(0);
            wsExcel.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1));
            wsExcel.GetRow(1).CreateCell(2).SetCellValue(fRPTName);
            wsExcel.AddMergedRegion(new CellRangeAddress(1, 1, 2, vColumnCount - 3));
            wsExcel.GetRow(1).GetCell(2).CellStyle = csTitle;
            wsExcel.GetRow(1).CreateCell(vColumnCount - 2).SetCellValue("印表日期：" + vPrintDate);
            wsExcel.AddMergedRegion(new CellRangeAddress(1, 1, vColumnCount - 2, vColumnCount - 1));
            wsExcel.CreateRow(2);
            wsExcel.GetRow(2).CreateCell(0).SetCellValue((int.Parse(eStopDate_Year.Text.Trim()) - 1911).ToString() + "年");
            //wsExcel.AutoSizeColumn(0);
            wsExcel.AddMergedRegion(new CellRangeAddress(2, 2, 0, vColumnCount - 1));
            wsExcel.GetRow(2).GetCell(0).CellStyle = csTitle;

            //建立報表內容
            //第一行是標題列
            wsExcel.CreateRow(3);
            for (int i = 0; i < gridAssetP005.Columns.Count; i++)
            {
                wsExcel.GetRow(3).CreateCell(i).SetCellValue(gridAssetP005.Columns[i].HeaderText);
            }
            //寫入內容
            for (int j = 0; j < gridAssetP005.Rows.Count; j++)
            {
                wsExcel.CreateRow(j + 4);
                for (int k = 0; k < gridAssetP005.Columns.Count; k++)
                {
                    if ((k >= 6) && (k <= 14)) //這幾個欄位是數值欄位
                    {
                        wsExcel.GetRow(j + 4).CreateCell(k);
                        wsExcel.GetRow(j + 4).GetCell(k).SetCellType(CellType.Numeric);
                        wsExcel.GetRow(j + 4).GetCell(k).SetCellValue(double.Parse(gridAssetP005.Rows[j].Cells[k].Text.Trim()));
                    }
                    else
                    {
                        wsExcel.GetRow(j + 4).CreateCell(k);
                        wsExcel.GetRow(j + 4).GetCell(k).SetCellType(CellType.String);
                        if (gridAssetP005.Columns[k].HeaderText == "取得日期")
                        {
                            vTempDate = DateTime.Parse(gridAssetP005.Rows[j].Cells[k].Text.Trim());
                            vGetDate = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                            wsExcel.GetRow(j + 4).GetCell(k).SetCellValue(vGetDate);
                        }
                        else
                        {
                            wsExcel.GetRow(j + 4).GetCell(k).SetCellValue(gridAssetP005.Rows[j].Cells[k].Text.Trim());
                        }
                    }

                }
            }
            //寫入小計欄位
            int TotalRows = gridAssetP005.Rows.Count + 4;
            wsExcel.CreateRow(TotalRows);
            wsExcel.GetRow(TotalRows).CreateCell(8).SetCellValue("合計：");
            wsExcel.GetRow(TotalRows).CreateCell(9).SetCellFormula("sum(J5:J" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(11).SetCellFormula("sum(L5:L" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(12).SetCellFormula("sum(M5:M" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(13).SetCellFormula("sum(N5:N" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(14).SetCellFormula("sum(O5:O" + TotalRows.ToString() + ")");

            try
            {
                /*
                MemoryStream msTarget = new MemoryStream();
                wbExcel.Write(msTarget);
                // 設定強制下載標頭
                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fSavePath + ".xls"));
                // 輸出檔案
                Response.BinaryWrite(msTarget.ToArray());
                msTarget.Close();

                Response.End();
                vResultStatus = true;//*/
                var msTarget = new NPOIMemoryStream();
                msTarget.AllowClose = false;
                wbExcel.Write(msTarget);
                msTarget.Flush();
                msTarget.Seek(0, SeekOrigin.Begin);
                msTarget.AllowClose = true;

                if (msTarget.Length > 0)
                {
                    //先判斷是不是IE
                    HttpContext.Current.Response.ContentType = "application/octet-stream";
                    HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                    string TourVerision = brObject.Type;
                    if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                    {
                        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + fSavePath + ".xls", System.Text.Encoding.UTF8));
                    }
                    else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                    {
                        // 設定強制下載標頭
                        Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fSavePath + ".xls"));
                    }
                    // 輸出檔案
                    Response.BinaryWrite(msTarget.ToArray());
                    msTarget.Close();

                    Response.End();
                    vResultStatus = true;
                }
            }
            catch (Exception eMessage)
            {
                vResultStatus = false;
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
            return vResultStatus;
        }

        private bool SaveToExcel_P006(string fRPTName, string fSavePath)
        {
            bool vResultStatus = false;
            int vColumnCount = gridAssetP006.Columns.Count;
            string vPrintDate = ((DateTime.Now.Year) - 1911).ToString("D3") + "/" + DateTime.Now.ToString("MM/dd");
            string vGetDate = "";
            DateTime vTempDate;
            //Excel 檔案
            HSSFWorkbook wbExcel = new HSSFWorkbook();
            //Excel 工作表
            HSSFSheet wsExcel = (HSSFSheet)wbExcel.CreateSheet(fRPTName);
            //設定標題欄位的格式
            HSSFCellStyle csTitle = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中
                                                            //設定標題欄位的字體格式
            HSSFFont fontTitle = (HSSFFont)wbExcel.CreateFont();
            //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
            fontTitle.IsBold = true;
            fontTitle.FontHeightInPoints = 20; //字體大小
            csTitle.SetFont(fontTitle);

            //建立報表表頭區
            wsExcel.CreateRow(0);
            wsExcel.GetRow(0).CreateCell(0).SetCellValue("桃園汽車客運股份有限公司");
            //wsExcel.AutoSizeColumn(0);
            wsExcel.AddMergedRegion(new CellRangeAddress(0, 0, 0, vColumnCount - 1));
            wsExcel.GetRow(0).GetCell(0).CellStyle = csTitle;
            wsExcel.CreateRow(1);
            wsExcel.GetRow(1).CreateCell(0).SetCellValue("");
            //wsExcel.AutoSizeColumn(0);
            wsExcel.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1));
            wsExcel.GetRow(1).CreateCell(2).SetCellValue(fRPTName);
            wsExcel.AddMergedRegion(new CellRangeAddress(1, 1, 2, vColumnCount - 3));
            wsExcel.GetRow(1).GetCell(2).CellStyle = csTitle;
            wsExcel.GetRow(1).CreateCell(vColumnCount - 2).SetCellValue("印表日期：" + vPrintDate);
            wsExcel.AddMergedRegion(new CellRangeAddress(1, 1, vColumnCount - 2, vColumnCount - 1));
            wsExcel.CreateRow(2);
            wsExcel.GetRow(2).CreateCell(0).SetCellValue((int.Parse(eStopDate_Year.Text.Trim()) - 1911).ToString() + "年");
            //wsExcel.AutoSizeColumn(0);
            wsExcel.AddMergedRegion(new CellRangeAddress(2, 2, 0, vColumnCount - 1));
            wsExcel.GetRow(2).GetCell(0).CellStyle = csTitle;

            //建立報表內容
            //第一行是標題列
            wsExcel.CreateRow(3);
            for (int i = 0; i < gridAssetP006.Columns.Count; i++)
            {
                wsExcel.GetRow(3).CreateCell(i).SetCellValue(gridAssetP006.Columns[i].HeaderText);
            }
            //寫入內容
            for (int j = 0; j < gridAssetP006.Rows.Count; j++)
            {
                wsExcel.CreateRow(j + 4);
                for (int k = 0; k < gridAssetP006.Columns.Count; k++)
                {
                    if ((k == 3) || ((k >= 5) && (k <= 8))) //這幾個欄位是數值欄位
                    {
                        wsExcel.GetRow(j + 4).CreateCell(k);
                        wsExcel.GetRow(j + 4).GetCell(k).SetCellType(CellType.Numeric);
                        wsExcel.GetRow(j + 4).GetCell(k).SetCellValue(double.Parse(gridAssetP006.Rows[j].Cells[k].Text.Trim()));
                    }
                    else
                    {
                        wsExcel.GetRow(j + 4).CreateCell(k);
                        wsExcel.GetRow(j + 4).GetCell(k).SetCellType(CellType.String);
                        if (gridAssetP006.Columns[k].HeaderText == "取得日期")
                        {
                            vTempDate = DateTime.Parse(gridAssetP006.Rows[j].Cells[k].Text.Trim());
                            vGetDate = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                            wsExcel.GetRow(j + 4).GetCell(k).SetCellValue(vGetDate);
                        }
                        else
                        {
                            wsExcel.GetRow(j + 4).GetCell(k).SetCellValue(gridAssetP006.Rows[j].Cells[k].Text.Trim());
                        }
                    }
                }
            }
            //寫入小計欄位
            int TotalRows = gridAssetP006.Rows.Count + 4;
            wsExcel.CreateRow(TotalRows);
            wsExcel.GetRow(TotalRows).CreateCell(2).SetCellValue("小計：");
            wsExcel.GetRow(TotalRows).CreateCell(3).SetCellFormula("sum(D5:D" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(5).SetCellFormula("sum(F5:F" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(6).SetCellFormula("sum(G5:G" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(7).SetCellFormula("sum(H5:H" + TotalRows.ToString() + ")");
            wsExcel.GetRow(TotalRows).CreateCell(8).SetCellFormula("sum(I5:I" + TotalRows.ToString() + ")");

            try
            {
                /*
                MemoryStream msTarget = new MemoryStream();
                wbExcel.Write(msTarget);
                // 設定強制下載標頭
                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fSavePath + ".xls"));
                // 輸出檔案
                Response.BinaryWrite(msTarget.ToArray());
                msTarget.Close();

                Response.End();
                vResultStatus = true;//*/
                var msTarget = new NPOIMemoryStream();
                msTarget.AllowClose = false;
                wbExcel.Write(msTarget);
                msTarget.Flush();
                msTarget.Seek(0, SeekOrigin.Begin);
                msTarget.AllowClose = true;

                if (msTarget.Length > 0)
                {
                    //先判斷是不是IE
                    HttpContext.Current.Response.ContentType = "application/octet-stream";
                    HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                    string TourVerision = brObject.Type;
                    if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                    {
                        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + fSavePath + ".xls", System.Text.Encoding.UTF8));
                    }
                    else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                    {
                        // 設定強制下載標頭
                        Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fSavePath + ".xls"));
                    }
                    // 輸出檔案
                    Response.BinaryWrite(msTarget.ToArray());
                    msTarget.Close();

                    Response.End();
                    vResultStatus = true;
                }
            }
            catch (Exception eMessage)
            {
                vResultStatus = false;
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
            return vResultStatus;
        }
    }
}