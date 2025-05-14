using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class CreateNameList : System.Web.UI.Page
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
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    string vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCalBaseDate_Search.ClientID;
                    string vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCalBaseDate_Search.Attributes["onClick"] = vDateScript;

                    vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eRetireHideDate_Search.ClientID;
                    vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eRetireHideDate_Search.Attributes["onClick"] = vDateScript;

                    if (!IsPostBack)
                    {
                        plReport.Visible = false;
                        eGiftYear.Text = (DateTime.Today.Year - 1911).ToString();
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

        private string GetSelectStr(string fSelectListType)
        {
            string vResultStr = "";
            string vGiftTitle = eGiftYear.Text.Trim() + " 年 " + rbListType.SelectedItem.Text + "發放名冊";
            DateTime vDate_S;
            string vDateStr_S;
            DateTime vDate_E;
            string vDateStr_E;
            DateTime vDate_Retire;
            string vDateStr_Retire;

            switch (fSelectListType)
            {
                case "0":
                case "1":
                case "2"://春節，端午，中秋
                    //先根據計算基準日算出起迄日期
                    vDateStr_E = (DateTime.TryParse(eCalBaseDate_Search.Text.Trim(), out vDate_E)) ? PF.GetMonthFirstDay(vDate_E, "B") : "";
                    vDateStr_S = (DateTime.TryParse(eCalBaseDate_Search.Text.Trim(), out vDate_S)) ? PF.GetMonthFirstDay(vDate_S.AddMonths(-13), "B") : "";
                    vDateStr_Retire = (DateTime.TryParse(eRetireHideDate_Search.Text.Trim(), out vDate_Retire)) ? PF.TransDateString(vDate_Retire, "B") : "";

                    if ((vDateStr_E.Trim() == "") || (vDateStr_S.Trim() == "") || (vDateStr_Retire.Trim() == ""))
                    {
                        //起迄日期其中之一沒有輸入
                        vResultStr = "";
                    }
                    else
                    {
                        vResultStr = "select * from (" + Environment.NewLine +
                                    "        select e.DepNo, d.[Name] as DepName, e.EmpNo, e.[Name], e.Title, db.ClassTxt as Title_C, e.Assumeday, e.WorkType, " + Environment.NewLine +
                                    "               case when isnull(pr.PayTime, 0) >= 6 then '全份' else '半份' end GiftType, " + Environment.NewLine +
                                    "               cast('' as nvarchar) GiftNote, cast('" + vGiftTitle.Trim() + "' as nvarchar) GiftTitle, " + Environment.NewLine +
                                    "               case when isnull(pr.PayTime, 0) >= 6 then 1 else 0 end FullGift, " + Environment.NewLine +
                                    "               case when isnull(pr.PayTime, 0) >= 6 then 0 else 1 end HalfGift, "+Environment.NewLine+
                                    "               cast('' as nvarchar) as StampPlace " +Environment.NewLine+
                                    "          from Employee as e left join Department as d on d.DepNo = e.DepNo " + Environment.NewLine +
                                    "                             left join DBDICB as db on db.ClassNo = e.Title and db.FKey = '員工資料        EMPLOYEE        TITLE' " + Environment.NewLine +
                                    "                             left join ( " + Environment.NewLine +
                                    "                                        select EmpNo, count(isnull(CashNum51, 0)) as PayTime " + Environment.NewLine +
                                    "                                          from PayRec " + Environment.NewLine +
                                    "                                         where PayDate between '" + vDateStr_S + "' and '" + vDateStr_E + "' " + Environment.NewLine +
                                    "                                           and PayDur = '1' " + Environment.NewLine +
                                    "                                         group by EmpNo " + Environment.NewLine +
                                    "                                       ) as pr on pr.EmpNo = e.EmpNo " + Environment.NewLine +
                                    "         where e.WorkType = '在職' " + Environment.NewLine +
                                    "           and ((isnull(e.DepNo, '00') <> '00') or ((isnull(e.DepNo, '00') = '00') and (e.Title  in ('010', '020', '030', '070')))) " + Environment.NewLine +
                                    "           and isnull(pr.PayTime, 0) > 0 " + Environment.NewLine +
                                    "         union all " + Environment.NewLine +
                                    "        select e2.DepNo, d2.[Name] as DepName, e2.EmpNo, e2.[Name], e2.Title, db2.ClassTxt as Title_C, e2.Assumeday, e2.WorkType, " + Environment.NewLine +
                                    "               cast('全份' as nvarchar) GiftType, cast('' as nvarchar) GiftNote, " + Environment.NewLine +
                                    "               cast('" + vGiftTitle.Trim() + "' as nvarchar) GiftTitle, 1 as FullGift, 0 as HalfGift, cast('' as nvarchar) as StampPlace " + Environment.NewLine +
                                    "          from Employee as e2 left join Department as d2 on d2.DepNo = e2.DepNo " + Environment.NewLine +
                                    "                              left join DBDICB as db2 on db2.ClassNo = e2.Title and db2.FKey = '員工資料        EMPLOYEE        TITLE' " + Environment.NewLine +
                                    "                              left join ( " + Environment.NewLine +
                                    "                                         select EmpNo, count(isnull(CashNum51, 0)) as PayTime " + Environment.NewLine +
                                    "                                           from PayRec " + Environment.NewLine +
                                    "                                          where PayDate between '" + vDateStr_S + "' and '" + vDateStr_E + "' " + Environment.NewLine +
                                    "                                            and PayDur = '1' " + Environment.NewLine +
                                    "                                          group by EmpNo " + Environment.NewLine +
                                    "                                        ) as pr2 on pr2.EmpNo = e2.EmpNo " + Environment.NewLine +
                                    "         where e2.WorkType = '退休' and e2.Leaveday >= '" + vDateStr_Retire + "' " + Environment.NewLine +
                                    "           and ((isnull(e2.DepNo, '00') <> '00') or ((isnull(e2.DepNo, '00') = '00') and (e2.Title  in ('010', '020', '030', '070')))) " + Environment.NewLine +
                                    ") z order by z.DepNo, z.Title, z.EmpNo";
                        /* 2025.04.30 規則全部重來
                        vDate_S = DateTime.Parse(eStartDate_Search.Text.Trim());
                        vDate_E = DateTime.Parse(eEndDate_Search.Text.Trim());
                        vResultStr = "select e.DepNo, d.[Name] as DepName, e.EmpNo, e.[Name], e.Title, db.ClassTxt as Title_C, e.Assumeday, e.WorkType, " + Environment.NewLine +
                                     "       case when ((e.WorkType = '退休') and (isnull(pr.PayTime, 0) > 0)) then '全份' " + Environment.NewLine +
                                     "            when((e.WorkType = '在職') and(isnull(pr.PayTime, 0) <= DateDiff(Month, '" + vDate_S.Year.ToString("D4") + "/" + vDate_S.ToString("MM/dd") + "', '" + vDate_E.Year.ToString("D4") + "/" + vDate_E.ToString("MM/dd") + "'))) then '半份' " + Environment.NewLine +
                                     "            else '全份' end GiftType, cast('' as nvarchar) GiftNote, cast('" + vGiftTitle.Trim() + "' as nvarchar) GiftTitle, " + Environment.NewLine +
                                     "       case when((e.WorkType = '退休') and(isnull(pr.PayTime, 0) > 0)) then 1 " + Environment.NewLine +
                                     "            when((e.WorkType = '在職') and(isnull(pr.PayTime, 0) <= DateDiff(Month, '" + vDate_S.Year.ToString("D4") + "/" + vDate_S.ToString("MM/dd") + "', '" + vDate_E.Year.ToString("D4") + "/" + vDate_E.ToString("MM/dd") + "'))) then 0 " + Environment.NewLine +
                                     "            else 1 end FullGift, " + Environment.NewLine +
                                     "       case when((e.WorkType = '退休') and(isnull(pr.PayTime, 0) > 0)) then 0 " + Environment.NewLine +
                                     "            when((e.WorkType = '在職') and(isnull(pr.PayTime, 0) <= DateDiff(Month, '" + vDate_S.Year.ToString("D4") + "/" + vDate_S.ToString("MM/dd") + "', '" + vDate_E.Year.ToString("D4") + "/" + vDate_E.ToString("MM/dd") + "'))) then 1 " + Environment.NewLine +
                                     "            else 0 end HalfGift, cast('' as nvarchar) as StampPlace " + Environment.NewLine +
                                     "  from Employee as e left join Department as d on d.DepNo = e.DepNo " + Environment.NewLine +
                                     "                     left join DBDICB as db on db.ClassNo = e.Title and db.FKey = '員工資料        EMPLOYEE        TITLE' " + Environment.NewLine +
                                     "                     left join ( " + Environment.NewLine +
                                     "                                select EmpNo, count(isnull(CashNum51, 0)) as PayTime " + Environment.NewLine +
                                     "                                  from PayRec " + Environment.NewLine +
                                     "                                 where PayDate between '" + vDate_S.Year.ToString("D4") + "/" + vDate_S.ToString("MM/dd") + "' and '" + vDate_E.Year.ToString("D4") + "/" + vDate_E.ToString("MM/dd") + "' " + Environment.NewLine +
                                     "                                   and PayDur = '1' " + Environment.NewLine +
                                     "                                 group by EmpNo " + Environment.NewLine +
                                     "                               ) as pr on pr.EmpNo = e.EmpNo " + Environment.NewLine +
                                     " where (e.WorkType = '在職' or e.WorkType = '退休') " + Environment.NewLine +
                                     "   and ((isnull(e.DepNo, '00') <> '00') or((isnull(e.DepNo, '00') = '00') and(e.Title  in ('010', '020', '030', '070')))) " + Environment.NewLine +
                                     "   and isnull(pr.PayTime, 0) > 0 " + Environment.NewLine +
                                     " order by e.DepNo, e.Title, e.EmpNo "; //*/
                        /* 2023.09.22 修改過濾標準
                        vResultStr = "select DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName, " + Environment.NewLine +
                                     "       EmpNo, [Name], " + Environment.NewLine +
                                     "       Title, (select ClassTxt from DBDICB where ClassNo = e.Title and FKey = '員工資料        EMPLOYEE        TITLE') Title_C, " + Environment.NewLine +
                                     "       AssumeDay, GiftType, cast('" + eNote.Text.Trim() + "' as nvarchar) GiftNote, " + Environment.NewLine +
                                     "       cast('" + vGiftTitle.Trim() + "' as nvarchar) GiftTitle, FullGift, HalfGift, cast('' as nvarchar) as StampPlace " + Environment.NewLine +
                                     "  from ( " + Environment.NewLine +
                                     "        select DepNo, EmpNo, [Name], Title, AssumeDay, cast('全份' as nvarchar) as GiftType, " + Environment.NewLine +
                                     "               cast(1 as integer) as FullGift, cast(0 as integer) as HalfGift " + Environment.NewLine +
                                     "          from Employee " + Environment.NewLine +
                                     "         where LeaveDay is null " + Environment.NewLine +
                                     "           and WorkType = '在職'" + Environment.NewLine +
                                     "           and isnull(DepNo, '00') <> '00' " + Environment.NewLine +
                                     "           and AssumeDay < '" + vDate_S.Year.ToString("D4") + "/" + vDate_S.ToString("MM/dd") + "' " + Environment.NewLine +
                                     "         union all " + Environment.NewLine +
                                     "        select DepNo, EmpNo, [Name], Title, AssumeDay, cast('全份' as nvarchar) as GiftType, " + Environment.NewLine +
                                     "               cast(1 as integer) as FullGift, cast(0 as integer) as HalfGift " + Environment.NewLine +
                                     "          from Employee " + Environment.NewLine +
                                     "         where LeaveDay is null " + Environment.NewLine +
                                     "           and WorkType = '在職'" + Environment.NewLine +
                                     "           and isnull(DepNo, '00') = '00' " + Environment.NewLine +
                                     "           and AssumeDay < '" + vDate_S.Year.ToString("D4") + "/" + vDate_S.ToString("MM/dd") + "' " + Environment.NewLine +
                                     "           and Title in ('010', '020', '030', '070') " + Environment.NewLine +
                                     "         union all " + Environment.NewLine +
                                     "        select DepNo, EmpNo, [Name], Title, AssumeDay, cast('半份' as nvarchar) as GiftType, " + Environment.NewLine +
                                     "               cast(0 as integer) as FullGift, cast(1 as integer) as HalfGift " + Environment.NewLine +
                                     "          from Employee " + Environment.NewLine +
                                     "         where LeaveDay is null " + Environment.NewLine +
                                     "           and WorkType = '在職'" + Environment.NewLine +
                                     "           and isnull(DepNo, '00') <> '00' " + Environment.NewLine +
                                     "           and AssumeDay >= '" + vDate_S.Year.ToString("D4") + "/" + vDate_S.ToString("MM/dd") + "' " + Environment.NewLine +
                                     "           and AssumeDay <= '" + vDate_E.Year.ToString("D4") + "/" + vDate_E.ToString("MM/dd") + "' " + Environment.NewLine +
                                     "         union all " + Environment.NewLine +
                                     "        select DepNo, EmpNo, [Name], Title, AssumeDay, cast('半份' as nvarchar) as GiftType, " + Environment.NewLine +
                                     "               cast(0 as integer) as FullGift, cast(1 as integer) as HalfGift " + Environment.NewLine +
                                     "          from Employee " + Environment.NewLine +
                                     "         where LeaveDay is null " + Environment.NewLine +
                                     "           and WorkType = '在職'" + Environment.NewLine +
                                     "           and isnull(DepNo, '00') = '00' " + Environment.NewLine +
                                     "           and AssumeDay >= '" + vDate_S.Year.ToString("D4") + "/" + vDate_S.ToString("MM/dd") + "' " + Environment.NewLine +
                                     "           and AssumeDay <= '" + vDate_E.Year.ToString("D4") + "/" + vDate_E.ToString("MM/dd") + "' " + Environment.NewLine +
                                     "           and Title in ('010', '020', '030', '070') " + Environment.NewLine +
                                     "       ) e " + Environment.NewLine +
                                     " where e.EmpNo <> 'supervisor' " + Environment.NewLine +
                                     " order by e.GiftType DESC, e.Title, e.DepNo, e.EmpNo"; //*/
                    }
                    break;
                case "3":
                    //紅包
                    if (eMoneyPay.Text.Trim() != "")
                    {
                        vResultStr = "select DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName, " + Environment.NewLine +
                                     "       EmpNo, [Name], " + Environment.NewLine +
                                     "       Title, (select ClassTxt from DBDICB where ClassNo = e.Title and FKey = '員工資料        EMPLOYEE        TITLE') Title_C, " + Environment.NewLine +
                                     "       AssumeDay, cast('" + eNote.Text.Trim() + "' as nvarchar) GiftNote, " + Environment.NewLine +
                                     "       cast(" + eMoneyPay.Text.Trim() + " as float) GiftPay, " + Environment.NewLine +
                                     "       cast('" + vGiftTitle.Trim() + "' as nvarchar) GiftTitle, cast('' as nvarchar) as StampPlace " + Environment.NewLine +
                                     "  from ( " + Environment.NewLine +
                                     "        select DepNo, EmpNo, [Name], Title, AssumeDay " + Environment.NewLine +
                                     "          from Employee " + Environment.NewLine +
                                     "         where LeaveDay is null " + Environment.NewLine +
                                     "           and WorkType = '在職'" + Environment.NewLine +
                                     "           and isnull(DepNo, '00') <> '00' " + Environment.NewLine +
                                     "         union all " + Environment.NewLine +
                                     "        select DepNo, EmpNo, [Name], Title, AssumeDay " + Environment.NewLine +
                                     "          from Employee " + Environment.NewLine +
                                     "         where LeaveDay is null " + Environment.NewLine +
                                     "           and WorkType = '在職'" + Environment.NewLine +
                                     "           and isnull(DepNo, '00') = '00' " + Environment.NewLine +
                                     "           and Title in ('010', '020', '030') " + Environment.NewLine +
                                     "       ) e " + Environment.NewLine +
                                     " where e.EmpNo <> 'supervisor' " + Environment.NewLine +
                                     " order by e.DepNo, e.Title, e.EmpNo";
                    }
                    else
                    {
                        vResultStr = "";
                    }
                    break;
                case "4"://尾牙
                    if (eMoneyPay.Text.Trim() != "")
                    {
                        vResultStr = "select DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName, " + Environment.NewLine +
                                     "       EmpNo, [Name], " + Environment.NewLine +
                                     "       Title, (select ClassTxt from DBDICB where ClassNo = e.Title and FKey = '員工資料        EMPLOYEE        TITLE') Title_C, " + Environment.NewLine +
                                     "       AssumeDay, cast('" + eNote.Text.Trim() + "' as nvarchar) GiftNote, " + Environment.NewLine +
                                     "       cast(" + eMoneyPay.Text.Trim() + " as float) GiftPay, " + Environment.NewLine +
                                     "       cast('" + vGiftTitle.Trim() + "' as nvarchar) GiftTitle, cast('' as nvarchar) as StampPlace " + Environment.NewLine +
                                     "  from ( " + Environment.NewLine +
                                     "        select DepNo, EmpNo, [Name], Title, AssumeDay " + Environment.NewLine +
                                     "          from Employee " + Environment.NewLine +
                                     "         where LeaveDay is null " + Environment.NewLine +
                                     "           and WorkType = '在職'" + Environment.NewLine +
                                     "           and isnull(DepNo, '00') <> '00' " + Environment.NewLine +
                                     "         union all " + Environment.NewLine +
                                     "        select DepNo, EmpNo, [Name], Title, AssumeDay " + Environment.NewLine +
                                     "          from Employee " + Environment.NewLine +
                                     "         where LeaveDay is null " + Environment.NewLine +
                                     "           and WorkType = '在職'" + Environment.NewLine +
                                     "           and isnull(DepNo, '00') = '00' " + Environment.NewLine +
                                     "           and Title in ('010', '020', '030', '070') " + Environment.NewLine +
                                     "       ) e " + Environment.NewLine +
                                     " where e.EmpNo <> 'supervisor' " + Environment.NewLine +
                                     " order by e.DepNo, e.Title, e.EmpNo";
                    }
                    else
                    {
                        vResultStr = "";
                    }
                    break;
                case "5"://月餅發放
                    vResultStr = "select DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName, " + Environment.NewLine +
                                 "       EmpNo, [Name], " + Environment.NewLine +
                                 "       Title, (select ClassTxt from DBDICB where ClassNo = e.Title and FKey = '員工資料        EMPLOYEE        TITLE') Title_C, " + Environment.NewLine +
                                 "       AssumeDay, cast('" + eNote.Text.Trim() + "' as nvarchar) GiftNote, " + Environment.NewLine +
                                 "       cast('" + vGiftTitle.Trim() + "' as nvarchar) GiftTitle, cast('' as nvarchar) as StampPlace " + Environment.NewLine +
                                 "  from ( " + Environment.NewLine +
                                 "        select DepNo, EmpNo, [Name], Title, AssumeDay " + Environment.NewLine +
                                 "          from Employee " + Environment.NewLine +
                                 "         where LeaveDay is null " + Environment.NewLine +
                                 "           and WorkType = '在職'" + Environment.NewLine +
                                 "           and isnull(DepNo, '00') <> '00' " + Environment.NewLine +
                                 "         union all " + Environment.NewLine +
                                 "        select DepNo, EmpNo, [Name], Title, AssumeDay " + Environment.NewLine +
                                 "          from Employee " + Environment.NewLine +
                                 "         where LeaveDay is null " + Environment.NewLine +
                                 "           and WorkType = '在職'" + Environment.NewLine +
                                 "           and isnull(DepNo, '00') = '00' " + Environment.NewLine +
                                 "           and Title in ('010', '020', '030', '070') " + Environment.NewLine +
                                 "       ) e " + Environment.NewLine +
                                 " where e.EmpNo <> 'supervisor' " + Environment.NewLine +
                                 " order by e.DepNo, e.Title, e.EmpNo";
                    break;
                case "6":
                    //制服發放
                    vResultStr = "select DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName, " + Environment.NewLine +
                                 "       EmpNo, [Name], " + Environment.NewLine +
                                 "       Title, (select ClassTxt from DBDICB where ClassNo = e.Title and FKey = '員工資料        EMPLOYEE        TITLE') Title_C, " + Environment.NewLine +
                                 "       AssumeDay, cast('" + eNote.Text.Trim() + "' as nvarchar) GiftNote, " + Environment.NewLine +
                                 "       cast('" + vGiftTitle.Trim() + "' as nvarchar) GiftTitle, cast('' as nvarchar) as StampPlace " + Environment.NewLine +
                                 "  from ( " + Environment.NewLine +
                                 "        select DepNo, EmpNo, [Name], Title, AssumeDay " + Environment.NewLine +
                                 "          from Employee " + Environment.NewLine +
                                 "         where LeaveDay is null " + Environment.NewLine +
                                 "           and WorkType = '在職'" + Environment.NewLine +
                                 "           and isnull(DepNo, '00') <> '00' " + Environment.NewLine +
                                 "         union all " + Environment.NewLine +
                                 "        select DepNo, EmpNo, [Name], Title, AssumeDay " + Environment.NewLine +
                                 "          from Employee " + Environment.NewLine +
                                 "         where LeaveDay is null " + Environment.NewLine +
                                 "           and WorkType = '在職'" + Environment.NewLine +
                                 "           and isnull(DepNo, '00') = '00' " + Environment.NewLine +
                                 "           and Title in ('010', '020', '030') " + Environment.NewLine +
                                 "       ) e " + Environment.NewLine +
                                 " where e.EmpNo <> 'supervisor' " + Environment.NewLine +
                                 " order by e.DepNo, e.Title, e.EmpNo";
                    break;
            }
            return vResultStr;
        }

        /// <summary>
        /// 三節禮金
        /// </summary>
        /// <param name="dtTemp"></param>
        private void Preview_Gift(DataTable dtTemp)
        {
            string vNoteStr = eGiftYear.Text.Trim() + " 年 " + rbListType.SelectedItem.Text + "發放名冊";
            ReportDataSource rdsPrint = new ReportDataSource("GiftNameListP_Gift", dtTemp);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\GiftNameListP_Gift.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.SetParameters(new ReportParameter("rpTitle", vNoteStr));
            rvPrint.LocalReport.SetParameters(new ReportParameter("rpNote", eNote.Text.Trim()));
            rvPrint.LocalReport.Refresh();
            plReport.Visible = true;
            plShowData.Visible = false;
            plSearch.Visible = false;
        }

        /// <summary>
        /// 制服發放
        /// </summary>
        /// <param name="dtTemp"></param>
        private void Preview_Uniform(DataTable dtTemp)
        {
            string vNoteStr = eGiftYear.Text.Trim() + " 年 " + rbListType.SelectedItem.Text + "發放名冊";
            ReportDataSource rdsPrint = new ReportDataSource("GiftNameListP_Uniform", dtTemp);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\GiftNameListP_Uniform.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.SetParameters(new ReportParameter("rpTitle", vNoteStr));
            rvPrint.LocalReport.SetParameters(new ReportParameter("rpNote", eNote.Text.Trim()));
            rvPrint.LocalReport.Refresh();
            plReport.Visible = true;
            plShowData.Visible = false;
            plSearch.Visible = false;
        }

        /// <summary>
        /// 紅包和尾牙
        /// </summary>
        /// <param name="dtTemp"></param>
        private void Preview_Money(DataTable dtTemp)
        {
            string vNoteStr = eGiftYear.Text.Trim() + " 年 " + rbListType.SelectedItem.Text + "發放名冊";
            ReportDataSource rdsPrint = new ReportDataSource("GiftNameListP_Money", dtTemp);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\GiftNameListP_Money.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.SetParameters(new ReportParameter("rpTitle", vNoteStr));
            rvPrint.LocalReport.SetParameters(new ReportParameter("rpNote", eNote.Text.Trim()));
            rvPrint.LocalReport.Refresh();
            plReport.Visible = true;
            plShowData.Visible = false;
            plSearch.Visible = false;
        }

        protected void rbListType_SelectedIndexChanged(object sender, EventArgs e)
        {
            eCalBaseDate_Search.Text = "";
            eRetireHideDate_Search.Text = "";
            switch (rbListType.SelectedValue)
            {
                case "0":
                case "1":
                case "2":
                    //春節，端午，中秋
                    eCalBaseDate_Search.Enabled = true;
                    eRetireHideDate_Search.Enabled = true;
                    eMoneyPay.Enabled = false;
                    break;
                case "3":
                case "4":
                    //紅包，尾牙
                    eCalBaseDate_Search.Enabled = false;
                    eRetireHideDate_Search.Enabled = false;
                    eMoneyPay.Enabled = true;
                    break;
                case "5":
                case "6":
                    //制服發放，月餅發放
                    eCalBaseDate_Search.Enabled = false;
                    eRetireHideDate_Search.Enabled = false;
                    eMoneyPay.Enabled = false;
                    break;
            }
        }

        /// <summary>
        /// 預覽報表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPreview_Click(object sender, EventArgs e)
        {
            string vSelectStr = "";

            switch (rbListType.SelectedValue)
            {
                case "0":
                case "1":
                case "2":
                    if (eCalBaseDate_Search.Text.Trim() == "")
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('基準日期不可空白！')");
                        Response.Write("</" + "Script>");
                        eCalBaseDate_Search.Focus();
                    }
                    else if (eRetireHideDate_Search.Text.Trim() == "")
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('退休起算日期不可空白！')");
                        Response.Write("</" + "Script>");
                        eRetireHideDate_Search.Focus();
                    }
                    else
                    {
                        vSelectStr = GetSelectStr(rbListType.SelectedValue);
                    }
                    break;
                case "3":
                case "4":
                    if (eMoneyPay.Text.Trim() == "")
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('發放金額不可空白！')");
                        Response.Write("</" + "Script>");
                        eMoneyPay.Focus();
                    }
                    else
                    {
                        vSelectStr = GetSelectStr(rbListType.SelectedValue);
                    }
                    break;
                default:
                    vSelectStr = GetSelectStr(rbListType.SelectedValue);
                    break;
            }

            if (vSelectStr != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                SqlConnection connNameList = new SqlConnection(vConnStr);
                SqlDataAdapter daNameList = new SqlDataAdapter(vSelectStr, connNameList);
                connNameList.Open();
                DataTable dtNameList = new DataTable();
                daNameList.Fill(dtNameList);
                switch (rbListType.SelectedValue)
                {
                    case "0":
                    case "1":
                    case "2":
                        Preview_Gift(dtNameList);
                        break;
                    case "3":
                    case "4":
                    case "5":
                        Preview_Money(dtNameList);
                        break;
                    case "6":
                        Preview_Uniform(dtNameList);
                        break;
                }
            }
        }

        /// <summary>
        /// 匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExportExcel_Click(object sender, EventArgs e)
        {
            DateTime vDate;
            string vSelectStr = "";

            switch (rbListType.SelectedValue)
            {
                case "0":
                case "1":
                case "2":
                    if (eCalBaseDate_Search.Text.Trim() == "")
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('基準日期不可空白！')");
                        Response.Write("</" + "Script>");
                        eCalBaseDate_Search.Focus();
                    }
                    else if (eRetireHideDate_Search.Text.Trim() == "")
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('退休起算日期不可空白！')");
                        Response.Write("</" + "Script>");
                        eRetireHideDate_Search.Focus();
                    }
                    else
                    {
                        vSelectStr = GetSelectStr(rbListType.SelectedValue);
                    }
                    break;
                case "3":
                case "4":
                    if (eMoneyPay.Text.Trim() == "")
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('發放金額不可空白！')");
                        Response.Write("</" + "Script>");
                        eMoneyPay.Focus();
                    }
                    else
                    {
                        vSelectStr = GetSelectStr(rbListType.SelectedValue);
                    }
                    break;
                default:
                    vSelectStr = GetSelectStr(rbListType.SelectedValue);
                    break;
            }


            if (vSelectStr != "")
            {
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
                fontTitle.FontHeightInPoints = 12; //字體大小
                csTitle.SetFont(fontTitle);

                //設定資料內容欄位的格式
                HSSFCellStyle csData = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                HSSFFont fontData = (HSSFFont)wbExcel.CreateFont();
                fontData.FontHeightInPoints = 12;
                csData.SetFont(fontData);

                HSSFCellStyle csData_Red = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                HSSFFont fontData_Red = (HSSFFont)wbExcel.CreateFont();
                fontData_Red.FontHeightInPoints = 12; //字體大小
                fontData_Red.Color = HSSFColor.Red.Index; //字體顏色
                csData_Red.SetFont(fontData_Red);

                HSSFCellStyle csData_Int = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                HSSFDataFormat format = (HSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Int.DataFormat = format.GetFormat("###,##0");

                string vHeaderText = "";
                int vLinesNo = 0;

                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                using (SqlConnection connExcel = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                    connExcel.Open();
                    SqlDataReader drExcel = cmdExcel.ExecuteReader();
                    if (drExcel.HasRows)
                    {
                        //查詢結果有資料的時候才執行
                        //新增一個工作表
                        wsExcel = (HSSFSheet)wbExcel.CreateSheet(rbListType.SelectedItem.Text.Trim());
                        //寫入標題列
                        vLinesNo = 0;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            vHeaderText = (drExcel.GetName(i) == "DepNo") ? "部門編號" :
                                          (drExcel.GetName(i) == "DepName") ? "部門" :
                                          (drExcel.GetName(i) == "EmpNo") ? "員工編號" :
                                          (drExcel.GetName(i) == "Name") ? "員工姓名" :
                                          (drExcel.GetName(i) == "Title") ? "職稱編號" :
                                          (drExcel.GetName(i) == "Title_C") ? "職稱" :
                                          (drExcel.GetName(i) == "AssumeDay") ? "到職日期" :
                                          (drExcel.GetName(i) == "GiftType") ? "份數" :
                                          (drExcel.GetName(i) == "GiftPay") ? "金額" :
                                          (drExcel.GetName(i) == "GiftNote") ? "發放說明" :
                                          (drExcel.GetName(i) == "GiftTitle") ? "發放主旨" :
                                          (drExcel.GetName(i) == "FullGift") ? "全份數量" :
                                          (drExcel.GetName(i) == "HalfGift") ? "半份數量" :
                                          (drExcel.GetName(i) == "StampPlace") ? "簽章" : drExcel.GetName(i);
                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                        }
                        vLinesNo++;
                        while (drExcel.Read())
                        {
                            wsExcel.CreateRow(vLinesNo);
                            for (int i = 0; i < drExcel.FieldCount; i++)
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(i);
                                if ((drExcel.GetName(i) == "GiftPay") || (drExcel.GetName(i) == "FullGift") || (drExcel.GetName(i) == "HalfGift"))
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                }
                                else if ((drExcel.GetName(i) == "AssumeDay") && (drExcel[i].ToString() != ""))
                                {
                                    vDate = DateTime.Parse(drExcel[i].ToString());
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue((vDate.Year - 1911).ToString("D3") + "/" + vDate.ToString("MM/dd"));
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                }
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            vLinesNo++;
                        }
                        try
                        {
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
                                string vFileName = rbListType.SelectedItem.Text.Trim();
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
                            Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                            Response.Write("</" + "Script>");
                            //throw;
                        }
                    }
                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbClosePreview_Click(object sender, EventArgs e)
        {
            plReport.Visible = false;
            plShowData.Visible = true;
            plSearch.Visible = true;
        }
    }
}