using Amaterasu_Function;
using System;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverWorkHoursStatistics : System.Web.UI.Page
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
                        eCalYear_Search.Text = (DateTime.Today.Year - 1911).ToString("D3");
                        eCalMonth_Search.Text = (DateTime.Today.Month).ToString("D2");
                        ClearData();
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

        private void ClearData()
        {
            fvDriverWorkHoursData.DataSourceID = "";
            plShowData.Visible = false;
        }

        protected void eDriverNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDriverNo = eDriverNo_Search.Text.Trim();
            string vDriverName = "";
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vDriverNo + "' and LeaveDay is null and Type = '20' ";
            vDriverName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDriverName == "")
            {
                vDriverName = vDriverNo;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' and LeaveDay is null and Type = '20' ";
                vDriverNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                if (vDriverNo == "")
                {
                    vDriverName = "";
                }
            }
            eDriverNo_Search.Text = vDriverNo;
            eDriverName_Search.Text = vDriverName;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            string vSQLStr = "";
            DateTime vThisMonth = DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today, "B"));
            DateTime vCalMonthDate = DateTime.Parse((int.Parse(eCalYear_Search.Text.Trim()) + 1911).ToString() + "/" + int.Parse(eCalMonth_Search.Text.Trim()).ToString("D2") + "/01");
            if (vCalMonthDate <= vThisMonth)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vCalYM = (Int32.Parse(eCalYear_Search.Text.Trim()) + 1911).ToString("D4") + "/" + Int32.Parse(eCalMonth_Search.Text.Trim()).ToString("D2");
                string vCalDate = vCalYM + "/01";
                string vDriverNo = eDriverNo_Search.Text.Trim();
                if (vDriverNo != "")
                {
                    if (vLoginDepNo != "09")
                    {
                        vSQLStr = "select count(EmpNo) RCount from Employee " + Environment.NewLine +
                                  " where EmpNo = '" + vDriverNo + "' and DepNo = '" + vLoginDepNo + "' ";
                    }
                    else
                    {
                        vSQLStr = "select count(EmpNo) RCount from Employee " + Environment.NewLine +
                                  " where EmpNo = '" + vDriverNo + "' ";
                    }
                    int vRCount = int.Parse(PF.GetValue(vConnStr, vSQLStr, "RCount"));
                    if (vRCount > 0)
                    {
                        /* //2019.04.10 變更資料來源，加入三種不同加班時數合計
                        vSQLStr = "select a.driver, (select [Name] from Employee where EmpNo = a.Driver) DriverName, '0' as NotOffDay ,'0' as IsShow , " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=1 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '1', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=2 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '2', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=3 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '3', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=4 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '4', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=5 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '5', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=6 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '6', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=7 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '7', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=8 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '8', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=9 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '9', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=10 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '10', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=11 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '11', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=12 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '12', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=13 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '13', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=14 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '14', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=15 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '15', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=16 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '16', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=17 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '17', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=18 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '18', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=19 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '19', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=20 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '20', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=21 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '21', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=22 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '22', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=23 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '23', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=24 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '24', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=25 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '25', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=26 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '26', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=27 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '27', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=28 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '28', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=29 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '29', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=30 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '30', " + Environment.NewLine +
                                  "       cast(isnull(sum(case when datepart(day,a.budate)=31 then a.workhr+a.workmin/100 else 0 end),0) as varchar(6)) as '31', " + Environment.NewLine +
                                  "       (sum(isnull(a.workhr,0)) + floor(sum(isnull(a.workmin,0)) / 60) + cast(cast(sum(isnull(a.workmin,0)) as int) % 60 as float) / 100 ) as totalhr, " + Environment.NewLine +
                                  "       sum(cast(isnull(a.rentnumber,'0') as float)) as rentnumber, " + Environment.NewLine +
                                  "       case when datediff(mm, '" + vCalDate + "', getdate()) = 0 then convert(int, day(Getdate())) else 32 end as day " + Environment.NewLine +
                                  "  from RUNSHEETA a " + Environment.NewLine +
                                  " where convert(varchar(7), a.budate, 111) = '" + vCalYM + "' " + Environment.NewLine +
                                  "   and a.driver = '" + vDriverNo + "' " + Environment.NewLine +
                                  " group by a.Driver "; //*/
                        vSQLStr = "select a.Driver, (select [Name] from Employee where EmpNo = a.Driver) DriverName, sum(cast(isnull(a.rentnumber,'0') as float)) as RentNumber, " + Environment.NewLine +
                                  "       (sum(isnull(a.workhr, 0)) + floor(sum(isnull(a.workmin, 0)) / 60) + cast(cast(sum(isnull(a.workmin, 0)) as int) % 60 as float) / 100) as TotalHR, " + Environment.NewLine +
                                  "       b.Hour01, b.Hour02, b.Hour03, b.Hour04, b.Hour05, b.Hour06, b.Hour07, b.Hour08, b.Hour09, b.Hour10, " + Environment.NewLine +
                                  "       b.Hour11, b.Hour12, b.Hour13, b.Hour14, b.Hour15, b.Hour16, b.Hour17, b.Hour18, b.Hour19, b.Hour20, " + Environment.NewLine +
                                  "       b.Hour21, b.Hour22, b.Hour23, b.Hour24, b.Hour25, b.Hour26, b.Hour27, b.Hour28, b.Hour29, b.Hour30, b.Hour31, " + Environment.NewLine +
                                  "       case when left(b.WorkState01, 1) = '1' then '休' when left(b.WorkState01, 1) = '2' then '例' when left(b.WorkState01, 1) = '3' then '國' else '' end as WorkState01, " + Environment.NewLine +
                                  "       case when left(b.WorkState02, 1) = '1' then '休' when left(b.WorkState02, 1) = '2' then '例' when left(b.WorkState02, 1) = '3' then '國' else '' end as WorkState02, " + Environment.NewLine +
                                  "       case when left(b.WorkState03, 1) = '1' then '休' when left(b.WorkState03, 1) = '2' then '例' when left(b.WorkState03, 1) = '3' then '國' else '' end as WorkState03, " + Environment.NewLine +
                                  "       case when left(b.WorkState04, 1) = '1' then '休' when left(b.WorkState04, 1) = '2' then '例' when left(b.WorkState04, 1) = '3' then '國' else '' end as WorkState04, " + Environment.NewLine +
                                  "       case when left(b.WorkState05, 1) = '1' then '休' when left(b.WorkState05, 1) = '2' then '例' when left(b.WorkState05, 1) = '3' then '國' else '' end as WorkState05, " + Environment.NewLine +
                                  "       case when left(b.WorkState06, 1) = '1' then '休' when left(b.WorkState06, 1) = '2' then '例' when left(b.WorkState06, 1) = '3' then '國' else '' end as WorkState06, " + Environment.NewLine +
                                  "       case when left(b.WorkState07, 1) = '1' then '休' when left(b.WorkState07, 1) = '2' then '例' when left(b.WorkState07, 1) = '3' then '國' else '' end as WorkState07, " + Environment.NewLine +
                                  "       case when left(b.WorkState08, 1) = '1' then '休' when left(b.WorkState08, 1) = '2' then '例' when left(b.WorkState08, 1) = '3' then '國' else '' end as WorkState08, " + Environment.NewLine +
                                  "       case when left(b.WorkState09, 1) = '1' then '休' when left(b.WorkState09, 1) = '2' then '例' when left(b.WorkState09, 1) = '3' then '國' else '' end as WorkState09, " + Environment.NewLine +
                                  "       case when left(b.WorkState10, 1) = '1' then '休' when left(b.WorkState10, 1) = '2' then '例' when left(b.WorkState10, 1) = '3' then '國' else '' end as WorkState10, " + Environment.NewLine +
                                  "       case when left(b.WorkState11, 1) = '1' then '休' when left(b.WorkState11, 1) = '2' then '例' when left(b.WorkState11, 1) = '3' then '國' else '' end as WorkState11, " + Environment.NewLine +
                                  "       case when left(b.WorkState12, 1) = '1' then '休' when left(b.WorkState12, 1) = '2' then '例' when left(b.WorkState12, 1) = '3' then '國' else '' end as WorkState12, " + Environment.NewLine +
                                  "       case when left(b.WorkState13, 1) = '1' then '休' when left(b.WorkState13, 1) = '2' then '例' when left(b.WorkState13, 1) = '3' then '國' else '' end as WorkState13, " + Environment.NewLine +
                                  "       case when left(b.WorkState14, 1) = '1' then '休' when left(b.WorkState14, 1) = '2' then '例' when left(b.WorkState14, 1) = '3' then '國' else '' end as WorkState14, " + Environment.NewLine +
                                  "       case when left(b.WorkState15, 1) = '1' then '休' when left(b.WorkState15, 1) = '2' then '例' when left(b.WorkState15, 1) = '3' then '國' else '' end as WorkState15, " + Environment.NewLine +
                                  "       case when left(b.WorkState16, 1) = '1' then '休' when left(b.WorkState16, 1) = '2' then '例' when left(b.WorkState16, 1) = '3' then '國' else '' end as WorkState16, " + Environment.NewLine +
                                  "       case when left(b.WorkState17, 1) = '1' then '休' when left(b.WorkState17, 1) = '2' then '例' when left(b.WorkState17, 1) = '3' then '國' else '' end as WorkState17, " + Environment.NewLine +
                                  "       case when left(b.WorkState18, 1) = '1' then '休' when left(b.WorkState18, 1) = '2' then '例' when left(b.WorkState18, 1) = '3' then '國' else '' end as WorkState18, " + Environment.NewLine +
                                  "       case when left(b.WorkState19, 1) = '1' then '休' when left(b.WorkState19, 1) = '2' then '例' when left(b.WorkState19, 1) = '3' then '國' else '' end as WorkState19, " + Environment.NewLine +
                                  "       case when left(b.WorkState20, 1) = '1' then '休' when left(b.WorkState20, 1) = '2' then '例' when left(b.WorkState20, 1) = '3' then '國' else '' end as WorkState20, " + Environment.NewLine +
                                  "       case when left(b.WorkState21, 1) = '1' then '休' when left(b.WorkState21, 1) = '2' then '例' when left(b.WorkState21, 1) = '3' then '國' else '' end as WorkState21, " + Environment.NewLine +
                                  "       case when left(b.WorkState22, 1) = '1' then '休' when left(b.WorkState22, 1) = '2' then '例' when left(b.WorkState22, 1) = '3' then '國' else '' end as WorkState22, " + Environment.NewLine +
                                  "       case when left(b.WorkState23, 1) = '1' then '休' when left(b.WorkState23, 1) = '2' then '例' when left(b.WorkState23, 1) = '3' then '國' else '' end as WorkState23, " + Environment.NewLine +
                                  "       case when left(b.WorkState24, 1) = '1' then '休' when left(b.WorkState24, 1) = '2' then '例' when left(b.WorkState24, 1) = '3' then '國' else '' end as WorkState24, " + Environment.NewLine +
                                  "       case when left(b.WorkState25, 1) = '1' then '休' when left(b.WorkState25, 1) = '2' then '例' when left(b.WorkState25, 1) = '3' then '國' else '' end as WorkState25, " + Environment.NewLine +
                                  "       case when left(b.WorkState26, 1) = '1' then '休' when left(b.WorkState26, 1) = '2' then '例' when left(b.WorkState26, 1) = '3' then '國' else '' end as WorkState26, " + Environment.NewLine +
                                  "       case when left(b.WorkState27, 1) = '1' then '休' when left(b.WorkState27, 1) = '2' then '例' when left(b.WorkState27, 1) = '3' then '國' else '' end as WorkState27, " + Environment.NewLine +
                                  "       case when left(b.WorkState28, 1) = '1' then '休' when left(b.WorkState28, 1) = '2' then '例' when left(b.WorkState28, 1) = '3' then '國' else '' end as WorkState28, " + Environment.NewLine +
                                  "       case when left(b.WorkState29, 1) = '1' then '休' when left(b.WorkState29, 1) = '2' then '例' when left(b.WorkState29, 1) = '3' then '國' else '' end as WorkState29, " + Environment.NewLine +
                                  "       case when left(b.WorkState30, 1) = '1' then '休' when left(b.WorkState30, 1) = '2' then '例' when left(b.WorkState30, 1) = '3' then '國' else '' end as WorkState30, " + Environment.NewLine +
                                  "       case when left(b.WorkState31, 1) = '1' then '休' when left(b.WorkState31, 1) = '2' then '例' when left(b.WorkState31, 1) = '3' then '國' else '' end as WorkState31, " + Environment.NewLine +
                                  "       b.CALHOURS_02, b.CALHOURS_03, b.CALHOURS_04 " + Environment.NewLine +
                                  "  from RunSheetA a left join DriverOverDuty b on b.OverYM = cast(year(a.budate) as varchar(4)) + right('00' + cast(month(a.BuDate) as varchar(2)), 2) and b.EmpNo = a.Driver " + Environment.NewLine +
                                  " where convert(varchar(7), a.budate, 111) = '" + vCalYM + "' and a.Driver = '" + vDriverNo + "' " + Environment.NewLine +
                                  " group by a.Driver, cast(year(a.budate) as varchar(4)) + right('00' + cast(month(a.BuDate) as varchar(2)), 2), " + Environment.NewLine +
                                  "       b.Hour01, b.Hour02, b.Hour03, b.Hour04, b.Hour05, b.Hour06, b.Hour07, b.Hour08, b.Hour09, b.Hour10, " + Environment.NewLine +
                                  "       b.Hour11, b.Hour12, b.Hour13, b.Hour14, b.Hour15, b.Hour16, b.Hour17, b.Hour18, b.Hour19, b.Hour20, " + Environment.NewLine +
                                  "       b.Hour21, b.Hour22, b.Hour23, b.Hour24, b.Hour25, b.Hour26, b.Hour27, b.Hour28, b.Hour29, b.Hour30, b.Hour31, " + Environment.NewLine +
                                  "       b.WorkState01, b.WorkState02, b.WorkState03, b.WorkState04, b.WorkState05, " + Environment.NewLine +
                                  "       b.WorkState06, b.WorkState07, b.WorkState08, b.WorkState09, b.WorkState10, " + Environment.NewLine +
                                  "       b.WorkState11, b.WorkState12, b.WorkState13, b.WorkState14, b.WorkState15, " + Environment.NewLine +
                                  "       b.WorkState16, b.WorkState17, b.WorkState18, b.WorkState19, b.WorkState20, " + Environment.NewLine +
                                  "       b.WorkState21, b.WorkState22, b.WorkState23, b.WorkState24, b.WorkState25, " + Environment.NewLine +
                                  "       b.WorkState26, b.WorkState27, b.WorkState28, b.WorkState29, b.WorkState30, b.WorkState31, " + Environment.NewLine +
                                  "       b.CALHOURS_02, b.CALHOURS_03, b.CALHOURS_04 ";
                        sdsDriverWorkHoursData.SelectCommand = "";
                        sdsDriverWorkHoursData.SelectCommand = vSQLStr;
                        fvDriverWorkHoursData.DataSourceID = "sdsDriverWorkHoursData";
                        sdsDriverWorkHoursData.DataBind();
                        plShowData.Visible = true;

                        string vRecordDriverStr = (eDriverNo_Search.Text.Trim() != "") ? eDriverNo_Search.Text.Trim() : "全部";
                        string vRecordNote = "查詢資料_駕駛員時數統計月結查詢" + Environment.NewLine +
                                             "DriverWorkHoursStatistics.aspx" + Environment.NewLine +
                                             "行車日期：" + eCalYear_Search.Text.Trim() + "年" + eCalMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                             "駕駛員：" + vRecordDriverStr;
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('在單位 [ " + Session["DepName"].ToString().Trim() + " ] 找不到駕駛員 [ " + vDriverNo + " ] 的資料')");
                        Response.Write("</" + "Script>");
                        ClearData();
                        eCalYear_Search.Focus();
                    }
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('不開放查詢未來的資料！')");
                Response.Write("</" + "Script>");
                ClearData();
                eCalYear_Search.Focus();
            }
        }

        protected void bbClear_Click(object sender, EventArgs e)
        {
            ClearData();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}