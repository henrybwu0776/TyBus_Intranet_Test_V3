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
using System.Threading;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class RunSheetCheckByStation : System.Web.UI.Page
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
        DateTime vToday = DateTime.Today;

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

                    string vDateURL_Temp = "";
                    string vDateScript_Temp = "";
                    if (!int.TryParse(vLoginDepNo, out int vDepIndex))
                    {
                        vDepIndex = -1;
                    }
                    //動態產生日期欄位的觸發事件
                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eBuDate_S_Search.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate_S_Search.Attributes["onClick"] = vDateScript_Temp;

                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eBuDate_E_Search.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate_E_Search.Attributes["onClick"] = vDateScript_Temp;
                    //
                    if (!IsPostBack)
                    {
                        eBuDate_S_Search.Text = PF.GetChinsesDate(DateTime.Today.AddDays(-2));
                        bbUpdateState.Visible = (vLoginDepNo == "09"); //2020.02.06 鎖定憑單的按鈕先只開放給電腦課
                        //2021.07.28 修改為只要是單位 01 ~ 10 都可以看多單位 
                        //if ((vLoginDepNo == "03") || (vLoginDepNo == "06") || (vLoginDepNo == "09"))
                        if ((vDepIndex > 0) && (vDepIndex < 11))
                        {
                            eDepNo_S.Text = "";
                            eDepName_S.Text = "";
                            eDepNo_S.Enabled = true;
                            eDepNo_E.Text = "";
                            eDepName_E.Text = "";
                            eDepNo_E.Enabled = true;
                        }
                        else
                        {
                            eDepNo_S.Text = vLoginDepNo;
                            eDepName_S.Text = (string)Session["DepName"];
                            eDepNo_S.Enabled = false;
                            eDepNo_E.Text = "";
                            eDepName_E.Text = "";
                            eDepNo_E.Enabled = false;
                        }
                        plSearch.Visible = true;
                        plShowData.Visible = false;
                        plPrint.Visible = false;
                    }
                    else
                    {
                        ListDataBind();
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
        /// 取回 SQL 查詢語法
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            DateTime vMonthFirstDate = (eBuDate_S_Search.Text.Trim() != "") ? DateTime.Parse(eBuDate_S_Search.Text.Trim()) : DateTime.Today.AddDays(-2);
            DateTime vMonthLastDate = (eBuDate_E_Search.Text.Trim() != "") ? DateTime.Parse(eBuDate_E_Search.Text.Trim()) : DateTime.Today.AddDays(-2);
            string vWStr_BuDate = ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() != "")) ?
                                  "         where BuDate between '" + PF.TransDateString(vMonthFirstDate, "B") + "' and '" + PF.TransDateString(vMonthLastDate, "B") + "' " + Environment.NewLine :
                                  ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() == "")) ?
                                  "         where BuDate = '" + PF.TransDateString(vMonthFirstDate, "B") + "' " + Environment.NewLine :
                                  ((eBuDate_S_Search.Text.Trim() == "") && (eBuDate_E_Search.Text.Trim() != "")) ?
                                  "         where BuDate = '" + PF.TransDateString(vMonthLastDate, "B") + "' " + Environment.NewLine :
                                  "         where BuDate = '" + PF.TransDateString(DateTime.Today.AddDays(-2), "B") + "' " + Environment.NewLine;
            string vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "           and a.DepNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? "           and a.DepNo = '" + eDepNo_S.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? "           and a.DepNo = '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vReturnStr = "select t.DepNo, (select [Name] from Department where DepNo = t.DepNo) DepName,  t.BuDate, t.AssignNo, t.Item, " + Environment.NewLine +
                                "       t.Driver,  (select [Name] from Employee where EmpNo = t.Driver) DriverName,  t.LinesNo, c.LineName, t.ErrorMSG " + Environment.NewLine +
                                "  from ( " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, cast('' as varchar) Item, a.Driver, " + Environment.NewLine +
                                "               cast(null as varchar(12)) LinesNo, cast('無上班時數' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (isnull(a.WorkHR, 0) * 60 + isnull(a.WorkMin, 0)) = 0 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, cast('' as varchar) Item, a.Driver, " + Environment.NewLine +
                                "               cast(null as varchar(12)) LinesNo, cast('上班分鐘數有誤' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and cast(isnull(a.WorkMin, '0') as int) % 10 <> 0 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, cast('' as varchar) Item, a.Driver, " + Environment.NewLine +
                                "               cast(null as varchar(12)) LinesNo, cast('租車加班費有誤' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and isnull(a.RentOverAMT, 0) <> 0 " + Environment.NewLine +
                                "           and isnull(a.RentOverAMT, 0) < 150 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, cast('' as varchar) Item, a.Driver, " + Environment.NewLine +
                                "               cast(null as varchar(12)) LinesNo, cast('駐在宿泊有誤' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and isnull(a.LiveAMT, 0) <> 0 " + Environment.NewLine +
                                "           and isnull(a.LiveAMT, 0) < 250 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('明細無公里數' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.CarType, '') <> '9' " + Environment.NewLine +
                                "           and isnull(b.ActualKM, 0) = 0 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('公里數有小數點第二位' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and len(cast((isnull(b.ActualKM, 0) -floor(isnull(b.ActualKM, 0))) as float)) -2 > 1 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('明細無車號' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.Car_ID, '') = '' " + Environment.NewLine +
                                "           and isnull(b.Item, '') <> '' " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('區間車/遊覽車明細公里數大於 1000 公里' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.CarType, '') in ('7', '8') " + Environment.NewLine +
                                "           and isnull(b.ActualKM, 0) > 1000 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('非區間車/遊覽車明細公里數大於 500 公里' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.CarType, '') not in ('7', '8') " + Environment.NewLine +
                                "           and isnull(b.ActualKM, 0) > 500 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('路線車種不同' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.CarType, '') <> (select l.CarType from Lines l where l.LinesNo = b.LinesNo)  " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('車種為交通車但明細沒有交通車收入' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.CarType, '') = '6' " + Environment.NewLine +
                                "           and isnull(b.ActualKM, 0) > 0 " + Environment.NewLine +
                                "           and isnull(b.unTraincome, 0) = 0 " + Environment.NewLine +
                                "           and (select isnull(Amount, 0) from Lines where LinesNo = b.LinesNo) > 0 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('車種為交通車也有交通車收入但沒有公里數' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.CarType, '') = '6' " + Environment.NewLine +
                                "           and isnull(b.unTraincome, 0) > 0 " + Environment.NewLine +
                                "           and isnull(b.ActualKM, 0) = 0 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('明細有交通車收入但車種不是交通車' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.CarType, '') <> '6' " + Environment.NewLine +
                                "           and isnull(b.unTraincome, 0) > 0 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('車種為租車但明細沒有租車收入' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.CarType, '') = '7' " + Environment.NewLine +
                                "           and isnull(b.ActualKM, 0) > 0 " + Environment.NewLine +
                                "           and isnull(b.RentIncomeA, 0) = 0 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('車種為租車也有租車收入但沒有公里數' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.CarType, '') = '7' " + Environment.NewLine +
                                "           and isnull(b.RentIncomeA, 0) > 0 " + Environment.NewLine +
                                "           and isnull(b.ActualKM, 0) = 0 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('明細有租車收入但車種不是租車' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.CarType, '') <> '7' " + Environment.NewLine +
                                "           and isnull(b.RentIncomeA, 0) > 0 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('車種為遊覽車但明細沒有遊覽車收入' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.CarType, '') = '8' " + Environment.NewLine +
                                "           and isnull(b.ActualKM, 0) > 0 " + Environment.NewLine +
                                "           and isnull(b.RentIncomeB, 0) = 0 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('車種為遊覽車也有遊覽車收入但沒有公里數' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.CarType, '') = '8' " + Environment.NewLine +
                                "           and isnull(b.RentIncomeB, 0) > 0 " + Environment.NewLine +
                                "           and isnull(b.ActualKM, 0) = 0 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('明細有遊覽車收入但車種不是遊覽車' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                "           and isnull(b.CarType, '') <> '8' " + Environment.NewLine +
                                "           and isnull(b.RentIncomeB, 0) > 0 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, cast('' as varchar) Item, Driver, " + Environment.NewLine +
                                "               cast(null as varchar(12)) LinesNo, cast('沒有租車趟次但是明細有區間租車或遊覽車代號' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and isnull(a.RentNumber, '0') = '0' " + Environment.NewLine +
                                "           and (select count(b.Item) RCount from RunSheetB b where b.AssignNo = a.AssignNo and isnull(b.CarType, '') in ('7', '8') and isnull(b.ReduceReason, '') = '') > 0 " + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, cast('' as varchar) Item, Driver, " + Environment.NewLine +
                                "               cast(null as varchar(12)) LinesNo, cast('有租車趟次但是明細沒有區間租車或遊覽車代號' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and isnull(a.RentNumber, '0') <> '0' " + Environment.NewLine +
                                "           and(select count(b.Item) RCount from RunSheetB b where b.AssignNo = a.AssignNo and isnull(b.CarType, '') in ('7', '8') and isnull(b.ReduceReason, '') = '') = 0 " + Environment.NewLine +
                                //2020.12.28 新增 "99908 異常"
                                "         union all " + Environment.NewLine +
                                "        select a.DepNo, a.BuDate, a.AssignNo, b.Item, Driver, " + Environment.NewLine +
                                "               b.LinesNo, cast('99908 去程或回程時間異常' as varchar(MAX)) ErrorMSG " + Environment.NewLine +
                                "          from RunSheetA a left join RunSheetB b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo +
                                "           and (select count(Item) from RunSheetB where AssignNo = a.AssignNo) > 0 " + Environment.NewLine +
                                "           and isnull(LTrim(RTrim(b.ReduceReason)), '') = '' " + Environment.NewLine +
                                "           and isnull(LTrim(RTrim(b.LinesNo)), '') = '99908' " + Environment.NewLine +
                                "           and ((isnull(LTrim(RTrim(b.ToTime)), ':') = ':') or " + Environment.NewLine +
                                "                (isnull(LTrim(RTrim(b.Totime)), ':') = '00:00') or" +
                                "                (isnull(LTrim(RTrim(b.BackTime)), ':') = ':') or " +
                                "                (isnull(LTrim(RTrim(b.BackTime)), ':') = '00:00' )) " + Environment.NewLine +
                                ") t left join Lines c on c.LinesNo = t.LinesNo " + Environment.NewLine +
                                " order by BuDate, DepNo";
            return vReturnStr;
        }

        private void ListDataBind()
        {
            string vSelectStr = GetSelectStr();
            sdsRSCheckResultList.SelectCommand = vSelectStr;
            gridRSCheckResultList.DataBind();
            plShowData.Visible = true;
        }

        private void SaveExcel(string fFileName, string fSelectStr)
        {
            DateTime vBudate;
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
            string vFieldValue = "";
            int vLinesNo = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(fSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(fFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "ASSIGNNO") ? "行車憑單單號" :
                                      (drExcel.GetName(i).ToUpper() == "BUDATE") ? "日期" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNO") ? "站別編號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "站別" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVER") ? "駕駛員工號" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVERNAME") ? "駕駛員" :
                                      (drExcel.GetName(i).ToUpper() == "ITEM") ? "憑單項次" :
                                      (drExcel.GetName(i).ToUpper() == "LINESNO") ? "路線" :
                                      (drExcel.GetName(i).ToUpper() == "LINENAME") ? "路線說明" :
                                      (drExcel.GetName(i).ToUpper() == "ERRORMSG") ? "錯誤資訊" : drExcel.GetName(i);
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    //以下開始寫入資料
                    vLinesNo++;
                    while (drExcel.Read())
                    {
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            if ((drExcel.GetName(i).ToUpper() == "ASSIGNNO") ||
                                (drExcel.GetName(i).ToUpper() == "DEPNO") ||
                                (drExcel.GetName(i).ToUpper() == "DEPNAME") ||
                                (drExcel.GetName(i).ToUpper() == "DRIVER") ||
                                (drExcel.GetName(i).ToUpper() == "DRIVERNAME") ||
                                (drExcel.GetName(i).ToUpper() == "ITEM") ||
                                (drExcel.GetName(i).ToUpper() == "LINESNO") ||
                                (drExcel.GetName(i).ToUpper() == "LINENAME") ||
                                (drExcel.GetName(i).ToUpper() == "ERRORMSG"))
                            {
                                //字串欄位
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                            }
                            else if ((drExcel.GetName(i).ToUpper() == "BUDATE") && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBudate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(((vBudate.Year) - 1911).ToString("D3") + "/" + vBudate.ToString("MM/dd"));
                            }
                            else
                            {
                                //數值欄位
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                vFieldValue = (drExcel[i].ToString().Trim() != "") ? drExcel[i].ToString().Trim() : "0";
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(vFieldValue));
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
                            string vDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "自 " + eDepNo_S.Text.Trim() + " 起至 " + eDepNo_E.Text.Trim() + " 止" :
                                               ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() :
                                               ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() : "全部";
                            string vBudateStr = ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() != "")) ? "自 " + eBuDate_S_Search.Text.Trim() + " 起至 " + eBuDate_E_Search.Text.Trim() + " 止" :
                                                ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() == "")) ? eBuDate_S_Search.Text.Trim() :
                                                ((eBuDate_S_Search.Text.Trim() == "") && (eBuDate_E_Search.Text.Trim() != "")) ? eBuDate_E_Search.Text.Trim() : "";
                            string vRecordNote = "匯出檔案_" + fFileName + Environment.NewLine +
                                                 "RunSheetCheckByStation.aspx" + Environment.NewLine +
                                                 "站別：" + vDepNoStr + Environment.NewLine +
                                                 "檢查日期：" + vBudateStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                            {
                                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + fFileName + ".xls", System.Text.Encoding.UTF8));
                            }
                            else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                            {
                                // 設定強制下載標頭
                                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fFileName + ".xls"));
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

        protected void eDepNo_S_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_S.Text.Trim();
            string vDepName_Temp = "";
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Departmwnt where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_S.Text = vDepNo_Temp;
            eDepName_S.Text = vDepName_Temp;
        }

        protected void eDepNo_E_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_E.Text.Trim();
            string vDepName_Temp = "";
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Departmwnt where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_E.Text = vDepNo_Temp;
            eDepName_E.Text = vDepName_Temp;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            ListDataBind();
        }

        /// <summary>
        /// 產出 EXCEL 檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vSelectStr = GetSelectStr();
            SaveExcel("各站行車記錄單異常查詢", vSelectStr);
        }

        /// <summary>
        /// 預覽報表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrint_Click(object sender, EventArgs e)
        {
            string vSelStr = GetSelectStr();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTempPrint = new SqlConnection(vConnStr))
            {
                string vDepNo_Temp = (Session["LoginDepNo"] != null) ? (String)Session["LoginDepNo"] : PF.GetValue(vConnStr, "select DepNo from Employee where EmpNo = '" + vLoginID + "' ", "DepNo");
                string vDepName_Temp = (Int32.Parse(vDepNo_Temp) >= 11) ? PF.GetValue(vConnStr, "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ", "Name") : "";
                SqlDataAdapter daPrintPoint = new SqlDataAdapter(vSelStr, connTempPrint);
                connTempPrint.Open();
                DataTable dtPrintPoint = new DataTable("RunSheetCheckByStationP");
                daPrintPoint.Fill(dtPrintPoint);
                if (dtPrintPoint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("RunSheetCheckByStationP", dtPrintPoint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\RunSheetCheckByStationP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("DepName", vDepName_Temp));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "行車記錄單異常"));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plShowData.Visible = false;
                    plSearch.Visible = false;

                    string vDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "自 " + eDepNo_S.Text.Trim() + " 起至 " + eDepNo_E.Text.Trim() + " 止" :
                                       ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() :
                                       ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() : "全部";
                    string vBudateStr = ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() != "")) ? "自 " + eBuDate_S_Search.Text.Trim() + " 起至 " + eBuDate_E_Search.Text.Trim() + " 止" :
                                        ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() == "")) ? eBuDate_S_Search.Text.Trim() :
                                        ((eBuDate_S_Search.Text.Trim() == "") && (eBuDate_E_Search.Text.Trim() != "")) ? eBuDate_E_Search.Text.Trim() : "";
                    string vRecordNote = "預覽報表_行車記錄單異常" + Environment.NewLine +
                                         "RunSheetCheckByStation.aspx" + Environment.NewLine +
                                         "站別：" + vDepNoStr + Environment.NewLine +
                                         "檢查日期：" + vBudateStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        /// <summary>
        /// 更新憑單狀態
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbUpdateState_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            DateTime vMonthFirstDate = (eBuDate_S_Search.Text.Trim() != "") ? DateTime.Parse(eBuDate_S_Search.Text.Trim()) : DateTime.Today.AddDays(-2);
            DateTime vMonthLastDate = (eBuDate_E_Search.Text.Trim() != "") ? DateTime.Parse(eBuDate_E_Search.Text.Trim()) : DateTime.Today.AddDays(-2);
            string vWStr_BuDate = ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() != "")) ?
                                  "   and BuDate between '" + PF.TransDateString(vMonthFirstDate, "B") + "' and '" + PF.TransDateString(vMonthLastDate, "B") + "' " + Environment.NewLine :
                                  ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() == "")) ?
                                  "   and BuDate = '" + PF.TransDateString(vMonthFirstDate, "B") + "' " + Environment.NewLine :
                                  ((eBuDate_S_Search.Text.Trim() == "") && (eBuDate_E_Search.Text.Trim() != "")) ?
                                  "   and BuDate = '" + PF.TransDateString(vMonthLastDate, "B") + "' " + Environment.NewLine :
                                  "   and BuDate = '" + PF.TransDateString(DateTime.Today.AddDays(-2), "B") + "' " + Environment.NewLine;
            string vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "   and DepNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? "   and DepNo = '" + eDepNo_S.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? "   and DepNo = '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vExecSQL = "update RunSheetA " + Environment.NewLine +
                              "   set IsInvestCheck = 'V', " + Environment.NewLine +
                              "       OnLineModifyDate = GetDate(), " + Environment.NewLine +
                              "       OnLineModifyMan = '" + vLoginID + "', " + Environment.NewLine +
                              "       OnLineModifyRemark = '線上系統批次審核' " + Environment.NewLine +
                              " where isnull(IsInvestCheck, '') <> 'V' " + Environment.NewLine +
                              vWStr_BuDate + vWStr_DepNo;
            try
            {
                PF.ExecSQL(vConnStr, vExecSQL);
            }
            catch (Exception eMessage)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plSearch.Visible = true;
            plShowData.Visible = true;
            plPrint.Visible = false;
        }
    }
}