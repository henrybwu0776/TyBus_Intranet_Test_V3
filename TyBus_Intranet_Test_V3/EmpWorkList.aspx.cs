using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
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
    public partial class EmpWorkList : Page
    {
        PublicFunction PF = new PublicFunction(); //加入公用程式碼參考
        private string vLoginID = "";
        private string vLoginName = "";
        private string vLoginDepNo = "";
        private string vLoginDepName = "";
        private string vLoginTitle = "";
        private string vLoginTitleName = "";
        private string vLoginEmpType = "";
        private string vConnStr = "";
        private string vComputerName = "";

        private DataTable dtTarget;

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
                vComputerName = Environment.MachineName; //2021.09.27 新增取得電腦名稱

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        eCalBase.Text = "6000";
                        eCalBase_Driver.Text = "10000";
                        plReport.Visible = false;
                        plCalculater.Visible = true;
                        plShow.Visible = true;
                    }
                }
                else
                {
                    Response.Redirect("~/Default.aspx");
                }
            }
            else
            {
                Response.Redirect("~/Default.aspx");
            }
        }

        /// <summary>
        /// 建立資料表結構
        /// </summary>
        /// <returns></returns>
        private DataTable CreateDataTable()
        {
            DataTable dtResult = new DataTable();

            DataColumn dcGetPayMode = new DataColumn("GetPayMode", typeof(String));
            dtResult.Columns.Add(dcGetPayMode);

            //DataColumn dcRowID = new DataColumn("RowID", typeof(Int32));
            //dtResult.Columns.Add(dcRowID);

            DataColumn dcIndexNo = new DataColumn("IndexNo", typeof(String));
            dtResult.Columns.Add(dcIndexNo);

            DataColumn dcDepNo = new DataColumn("DepNo", typeof(String));
            dtResult.Columns.Add(dcDepNo);

            DataColumn dcDepName = new DataColumn("DepName", typeof(String));
            dtResult.Columns.Add(dcDepName);

            DataColumn dcEmpNo = new DataColumn("EmpNo", typeof(String));
            dtResult.Columns.Add(dcEmpNo);

            DataColumn dcEmpName = new DataColumn("EmpName", typeof(String));
            dtResult.Columns.Add(dcEmpName);

            DataColumn dcIDCardNo = new DataColumn("IDCardNo", typeof(String));
            dtResult.Columns.Add(dcIDCardNo);

            DataColumn dcTitle = new DataColumn("Title", typeof(String));
            dtResult.Columns.Add(dcTitle);

            DataColumn dcMonthDays = new DataColumn("MonthDays", typeof(Double));
            dtResult.Columns.Add(dcMonthDays);

            DataColumn dcBaseAmount = new DataColumn("BaseAmount", typeof(Double));
            dtResult.Columns.Add(dcBaseAmount);

            DataColumn dcHolidays = new DataColumn("Holidays", typeof(Double));
            dtResult.Columns.Add(dcHolidays);

            DataColumn dcESCDays = new DataColumn("ESCDays", typeof(Double));
            dtResult.Columns.Add(dcESCDays);

            DataColumn dcWorkDays = new DataColumn("WorkDays", typeof(Double));
            dtResult.Columns.Add(dcWorkDays);

            DataColumn dcBoundsRatio = new DataColumn("BoundsRatio", typeof(Double));
            dtResult.Columns.Add(dcBoundsRatio);

            DataColumn dcBounds = new DataColumn("Bounds", typeof(Double));
            dtResult.Columns.Add(dcBounds);

            DataColumn dcWorkType = new DataColumn("WorkType", typeof(String));
            dtResult.Columns.Add(dcWorkType);

            DataColumn dcRemark = new DataColumn("Remark", typeof(String));
            dtResult.Columns.Add(dcRemark);

            return dtResult;
        }

        private string GetSelectStr()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vCalDate = eCalYM.Text.Trim().Substring(0, 4) + "/" + eCalYM.Text.Trim().Substring(4, 2) + "/01";
            string vCalLastDay = PF.GetMonthLastDay(DateTime.Parse(vCalDate), "B");
            string vCalDays = PF.GetMonthDays(DateTime.Parse(vCalDate)).ToString();
            string vSelectStr = "";
            /* 2024.04.30 改成不管當月在職幾天都給全額
            string vSelectStr = "select ROW_NUMBER() OVER(PARTITION BY e.DepNo  order by e.DepNo, e.Title, e.EmpNo) RowID, " + Environment.NewLine +
                                "       e.DepNo, d.[Name] DepName, e.EmpNo, e.[Name] EmpName, e.Title, a.ClassTxt, isnull(e.Leaveday, '') Leaveday, " + Environment.NewLine +
                                "       case when isnull(e.LeaveDay, '') = '' then " + vCalDays + Environment.NewLine +
                                "            when e.Leaveday > '" + vCalLastDay + "' then " + vCalDays + Environment.NewLine +
                                "            when e.LeaveDay between '" + vCalDate + "' and '" + vCalLastDay + "' then DateDiff(day, '" + vCalDate + "', e.LeaveDay) + 1 end WorkDays " + Environment.NewLine +
                                "  from Employee e left join Department d on d.DepNo = e.DepNo " + Environment.NewLine +
                                "                  left join DBDICB a on a.ClassNo = e.Title and a.Fkey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                                " where e.AssumeDay < '" + vCalDate + "' " + Environment.NewLine +
                                "   and (isnull(e.LeaveDay, '') = '' or e.LeaveDay >= '" + vCalDate + "') " + Environment.NewLine +
                                "   and (e.DepNo between '01' and '10' or(e.DepNo > '11' and e.Type != '20')) "; //*/
            /* 2024.07.04 修改條件，須當月全月在職人員才給
            string vSelectStr = "select ROW_NUMBER() OVER(PARTITION BY e.DepNo  order by e.DepNo, e.Title, e.EmpNo) RowID, e.IDCardNo, e.WorkType, " + Environment.NewLine +
                                "       e.DepNo, d.[Name] DepName, e.EmpNo, e.[Name] EmpName, e.Title, a.ClassTxt, isnull(e.Leaveday, '') Leaveday, " + Environment.NewLine +
                                "       case when isnull(e.LeaveDay, '') = '' then " + vCalDays + Environment.NewLine +
                                "            when e.Leaveday > '" + vCalLastDay + "' then " + vCalDays + Environment.NewLine +
                                "            when e.LeaveDay between '" + vCalDate + "' and '" + vCalLastDay + "' then DateDiff(day, '" + vCalDate + "', e.LeaveDay) + 1 end WorkDays " + Environment.NewLine +
                                "  from Employee e left join Department d on d.DepNo = e.DepNo " + Environment.NewLine +
                                "                  left join DBDICB a on a.ClassNo = e.Title and a.Fkey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                                " where e.AssumeDay <= '" + vCalLastDay + " 23:59:59' " + Environment.NewLine +
                                "   and (isnull(e.LeaveDay, '') = '' or e.LeaveDay >= '" + vCalDate + " 00:00:00') " + Environment.NewLine +
                                "   and e.DepNo > '00' " + Environment.NewLine +
                                "   and e.Title != '271' "; //*/
            /* 2024.07.09 新增未發放名單
            vSelectStr = "select ROW_NUMBER() OVER(PARTITION BY e.DepNo  order by e.DepNo, e.Title, e.EmpNo) RowID, e.IDCardNo, e.WorkType, " + Environment.NewLine +
                         "       e.DepNo, d.[Name] DepName, e.EmpNo, e.[Name] EmpName, e.Title, a.ClassTxt, isnull(e.Leaveday, '') Leaveday, " + Environment.NewLine +
                         "       " + vCalDays + " as WorkDays " + Environment.NewLine +
                         "  from Employee as e left join Department as d on d.DepNo = e.DepNo " + Environment.NewLine +
                         "                     left join DBDICB as a on a.ClassNo = e.Title and a.FKey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                         " where e.Title != '271' " + Environment.NewLine +
                         "   and e.DepNo > '00' " + Environment.NewLine +
                         "   and e.Assumeday <= '" + vCalDate + " 23:59:59' " + Environment.NewLine +
                         "   and (isnull(e.LeaveDay, '') = '' or e.LeaveDay > '" + vCalLastDay + " 00:00:00' )" + Environment.NewLine +
                         "   and (isnull(e.BeginDay, '') = '' " + Environment.NewLine +
                         "       or (isnull(e.BeginDay, '') > '" + vCalLastDay + "') " + Environment.NewLine +
                         "       or (isnull(e.StopDay, '') < '" + vCalDate + "')) " + Environment.NewLine +
                         "   and (select count(ESCType) from ESCDuty where ApplyMan = e.EmpNo and RealDay between '" + vCalDate + " 00:00:00' and '" + vCalLastDay + " 23:59:59' and ESCType in ('03', '13')) = 0"; //*/
            vSelectStr = "select Distinct * from ( " + Environment.NewLine +
                         "       select '1' as GetPayMode, e.IDCardNo, e.WorkType, e.DepNo, d.[Name] DepName, e.EmpNo, e.[Name] EmpName, e.Title, " + Environment.NewLine +
                         "              a.ClassTxt, isnull(e.Leaveday, '') Leaveday, NULL as WorkDays, NULL as Remark " + Environment.NewLine +
                         "         from Employee as e left join Department as d on d.DepNo = e.DepNo " + Environment.NewLine +
                         "                            left join DBDICB as a on a.ClassNo = e.Title and a.FKey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                         "        where e.Title != '271' " + Environment.NewLine +
                         "          and e.DepNo > '00' " + Environment.NewLine +
                         "          and e.Assumeday <= '" + vCalDate + " 23:59:59' " + Environment.NewLine +
                         "          and (isnull(e.LeaveDay, '') = '' or e.LeaveDay > '" + vCalLastDay + " 00:00:00') " + Environment.NewLine +
                         "          and (isnull(e.BeginDay, '') = '' " + Environment.NewLine +
                         "              or (isnull(e.BeginDay, '') > '" + vCalLastDay + "') " + Environment.NewLine +
                         "              or (isnull(e.StopDay, '" + vCalLastDay + "') < '" + vCalDate + "')) " + Environment.NewLine +
                         "          and (select count(ESCType) from ESCDuty where ApplyMan = e.EmpNo and RealDay between '" + vCalDate + " 00:00:00' and '" + vCalLastDay + " 23:59:59' and ESCType in ('03', '13')) = 0 " + Environment.NewLine +
                         "          and (select count(AssignNo) RCount from RunSheetB where LinesNo in ('952', '03560') and AssignNo in (select AssignNo from RunSheetA where Driver = e.EmpNo and BuDate between '" + vCalDate + "' and '" + vCalLastDay + "')) = 0 " + Environment.NewLine +
                         "        union all " + Environment.NewLine +
                         "       select '0' as GetPayMode, e.IDCardNo, e.WorkType, e.DepNo, d.[Name] DepName, e.EmpNo, e.[Name] EmpName, e.Title, " + Environment.NewLine +
                         "              a.ClassTxt, isnull(e.Leaveday, '') Leaveday, NULL as WorkDays, 'MOU員工' as Remark " + Environment.NewLine +
                         "         from Employee as e left join Department as d on d.DepNo = e.DepNo " + Environment.NewLine +
                         "                            left join DBDICB as a on a.ClassNo = e.Title and a.FKey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                         "        where e.Title = '271' " + Environment.NewLine +
                         "          and e.DepNo > '00' " + Environment.NewLine +
                         //2024.09.12 MOU 人員區分 "正常在職、本月到職、本月離職
                         "          and e.Assumeday < '"+vCalDate+" 00:00:00' "+Environment.NewLine+
                         "          and isnull(e.Leaveday, '') = '' "+Environment.NewLine +
                         "        union all " + Environment.NewLine +
                         "       select '0' as GetPayMode, e.IDCardNo, e.WorkType, e.DepNo, d.[Name] DepName, e.EmpNo, e.[Name] EmpName, e.Title, " + Environment.NewLine +
                         "              a.ClassTxt, isnull(e.Leaveday, '') Leaveday, NULL as WorkDays, 'MOU員工_本月到職' as Remark " + Environment.NewLine +
                         "         from Employee as e left join Department as d on d.DepNo = e.DepNo " + Environment.NewLine +
                         "                            left join DBDICB as a on a.ClassNo = e.Title and a.FKey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                         "        where e.Title = '271' " + Environment.NewLine +
                         "          and e.DepNo > '00' " + Environment.NewLine +
                         "          and e.Assumeday > '" + vCalDate + " 23:59:59' " + Environment.NewLine +
                         "        union all " + Environment.NewLine +
                         "       select '0' as GetPayMode, e.IDCardNo, e.WorkType, e.DepNo, d.[Name] DepName, e.EmpNo, e.[Name] EmpName, e.Title, " + Environment.NewLine +
                         "              a.ClassTxt, isnull(e.Leaveday, '') Leaveday, NULL as WorkDays, 'MOU員工_本月離職' as Remark " + Environment.NewLine +
                         "         from Employee as e left join Department as d on d.DepNo = e.DepNo " + Environment.NewLine +
                         "                            left join DBDICB as a on a.ClassNo = e.Title and a.FKey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                         "        where e.Title = '271' " + Environment.NewLine +
                         "          and e.DepNo > '00' " + Environment.NewLine +
                         "          and isnull(e.LeaveDay, '') > '" + vCalDate + " 00:00:00' " + Environment.NewLine +
                         //===============================================================================================================
                         "        union all " + Environment.NewLine +
                         "       select '0' as GetPayMode, e.IDCardNo, e.WorkType, e.DepNo, d.[Name] DepName, e.EmpNo, e.[Name] EmpName, e.Title, " + Environment.NewLine +
                         "              a.ClassTxt, isnull(e.Leaveday, '') Leaveday, NULL as WorkDays, '本月到職' as Remark " + Environment.NewLine +
                         "         from Employee as e left join Department as d on d.DepNo = e.DepNo " + Environment.NewLine +
                         "                            left join DBDICB as a on a.ClassNo = e.Title and a.FKey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                         "        where e.Title != '271' " + Environment.NewLine +
                         "          and e.DepNo > '00' " + Environment.NewLine +
                         "          and e.Assumeday > '" + vCalDate + " 23:59:59' " + Environment.NewLine +
                         "        union all " + Environment.NewLine +
                         "       select '0' as GetPayMode, e.IDCardNo, e.WorkType, e.DepNo, d.[Name] DepName, e.EmpNo, e.[Name] EmpName, e.Title, " + Environment.NewLine +
                         "              a.ClassTxt, isnull(e.Leaveday, '') Leaveday, NULL as WorkDays, '本月離職' as Remark " + Environment.NewLine +
                         "         from Employee as e left join Department as d on d.DepNo = e.DepNo " + Environment.NewLine +
                         "                            left join DBDICB as a on a.ClassNo = e.Title and a.FKey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                         "        where e.Title != '271' " + Environment.NewLine +
                         "          and e.DepNo > '00' " + Environment.NewLine +
                         "          and isnull(e.LeaveDay, '') > '" + vCalDate + " 00:00:00' " + Environment.NewLine +
                         "        union all " + Environment.NewLine +
                         "       select '0' as GetPayMode, e.IDCardNo, e.WorkType, e.DepNo, d.[Name] DepName, e.EmpNo, e.[Name] EmpName, e.Title, " + Environment.NewLine +
                         "              a.ClassTxt, isnull(e.Leaveday, '') Leaveday, NULL as WorkDays, '有留職停薪' as Remark " + Environment.NewLine +
                         "         from Employee as e left join Department as d on d.DepNo = e.DepNo " + Environment.NewLine +
                         "                            left join DBDICB as a on a.ClassNo = e.Title and a.FKey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                         "        where e.Title != '271' " + Environment.NewLine +
                         "          and e.DepNo > '00' " + Environment.NewLine +
                         "          and ((isnull(e.BeginDay, '') != '') and(isnull(e.StopDay, '') = '') " + Environment.NewLine +
                         "              or (isnull(e.BeginDay, '') between '" + vCalDate + "' and '" + vCalLastDay + "') " + Environment.NewLine +
                         "              or (isnull(e.StopDay, '') between '" + vCalDate + "' and '" + vCalLastDay + "')) " + Environment.NewLine +
                         "        union all " + Environment.NewLine +
                         "       select '0' as GetPayMode, e.IDCardNo, e.WorkType, e.DepNo, d.[Name] DepName, e.EmpNo, e.[Name] EmpName, e.Title, " + Environment.NewLine +
                         "              a.ClassTxt, isnull(e.Leaveday, '') Leaveday, NULL as WorkDays, '有公傷假' as Remark " + Environment.NewLine +
                         "         from Employee as e left join Department as d on d.DepNo = e.DepNo " + Environment.NewLine +
                         "                            left join DBDICB as a on a.ClassNo = e.Title and a.FKey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                         "        where e.Title != '271' " + Environment.NewLine +
                         "          and e.DepNo > '00' " + Environment.NewLine +
                         "          and (select count(ESCType) from ESCDuty where ApplyMan = e.EmpNo and RealDay between '" + vCalDate + " 00:00:00' and '" + vCalLastDay + " 23:59:59' and ESCType = '03') > 0 " + Environment.NewLine +
                         "        union all " + Environment.NewLine +
                         "       select '0' as GetPayMode, e.IDCardNo, e.WorkType, e.DepNo, d.[Name] DepName, e.EmpNo, e.[Name] EmpName, e.Title, " + Environment.NewLine +
                         "              a.ClassTxt, isnull(e.Leaveday, '') Leaveday, NULL as WorkDays, '有曠職' as Remark " + Environment.NewLine +
                         "         from Employee as e left join Department as d on d.DepNo = e.DepNo " + Environment.NewLine +
                         "                            left join DBDICB as a on a.ClassNo = e.Title and a.FKey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                         "        where e.Title != '271' " + Environment.NewLine +
                         "          and e.DepNo > '00' " + Environment.NewLine +
                         "          and (select count(ESCType) from ESCDuty where ApplyMan = e.EmpNo and RealDay between '" + vCalDate + " 00:00:00' and '" + vCalLastDay + " 23:59:59' and ESCType = '13') > 0 " + Environment.NewLine +
                         "        union all " + Environment.NewLine +
                         "       select '0' as GetPayMode, e.IDCardNo, e.WorkType, e.DepNo, d.[Name] DepName, e.EmpNo, e.[Name] EmpName, e.Title, " + Environment.NewLine +
                         "              a.ClassTxt, isnull(e.Leaveday, '') Leaveday, NULL as WorkDays, '952 路線駕駛' as Remark " + Environment.NewLine +
                         "         from Employee as e left join Department as d on d.DepNo = e.DepNo " + Environment.NewLine +
                         "                            left join DBDICB as a on a.ClassNo = e.Title and a.FKey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                         "        where (select count(AssignNo) RCount from RunSheetB where LinesNo in ('952', '03560') and AssignNo in (select AssignNo from RunSheetA where Driver = e.EmpNo and BuDate between '" + vCalDate + "' and '" + vCalLastDay + "')) > 0 " + Environment.NewLine +
                         ") as t order by t.GetPayMode DESC, t.DepNo, t.Title, t.EmpNo";
            return vSelectStr;
        }

        private void CalData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vGetEmpList = GetSelectStr();
            string vCalYM = eCalYM.Text.Trim();
            string vBaseAmount = eCalBase.Text.Trim();
            string vBaseAmount_D = eCalBase_Driver.Text.Trim();
            string vCalDate = eCalYM.Text.Trim().Substring(0, 4) + "/" + eCalYM.Text.Trim().Substring(4, 2) + "/01";
            string vCalLastDay = PF.GetMonthLastDay(DateTime.Parse(vCalDate), "B");
            string vCalDays = PF.GetMonthDays(DateTime.Parse(vCalDate)).ToString();
            string vEmpNo = "";
            string vDepNo = "";
            string vHoliday = "";
            string vESCDays = "";
            string vTempStr = "";
            double vCalDays_D = 0.0;
            double vHolidays_D = 0.0;
            double vESCDays_D = 0.0;
            double vBoundsRatio_D = 0.0;
            double vBounds_D = 0.0;
            string vWorkType = "";

            using (SqlConnection connGetEmpList = new SqlConnection(vConnStr))
            {
                SqlCommand cmdGetEmpList = new SqlCommand(vGetEmpList, connGetEmpList);
                connGetEmpList.Open();
                SqlDataReader drGetEmpList = cmdGetEmpList.ExecuteReader();
                if (drGetEmpList.HasRows)
                {
                    dtTarget = CreateDataTable();
                }
                while (drGetEmpList.Read())
                {
                    vEmpNo = drGetEmpList["EmpNo"].ToString().Trim();
                    vWorkType = drGetEmpList["WorkType"].ToString().Trim();
                    vDepNo = drGetEmpList["DepNo"].ToString().Trim();
                    vTempStr = (Int32.Parse(vDepNo) <= 10) ?
                               "select cast(Holiday as varchar) Holiday from BusHoliday where BuDate = '" + vCalDate + "' and Kind = '1' " :
                               "select cast(Holiday as varchar) Holiday from BusHoliday where BuDate = '" + vCalDate + "' and Kind = '2' ";
                    vHoliday = PF.GetValue(vConnStr, vTempStr, "Holiday");
                    vTempStr = "select sum(Hours) / 8.0 ESCDays from ESCDuty " + Environment.NewLine +
                               " where RealDay Between '" + vCalDate + "' and '" + vCalLastDay + "' " + Environment.NewLine +
                               "   and ESCType not in ('01', '03') " + Environment.NewLine +
                               "   and ApplyMan = '" + vEmpNo + "' ";
                    vESCDays = PF.GetValue(vConnStr, vTempStr, "ESCDays");
                    DataRow rowTemp = dtTarget.NewRow();
                    rowTemp["GetPayMode"] = drGetEmpList["GetPayMode"];
                    rowTemp["IndexNo"] = vCalYM + vEmpNo;
                    rowTemp["DepNo"] = vDepNo;
                    rowTemp["DepName"] = drGetEmpList["DepName"];
                    rowTemp["EmpNo"] = vEmpNo;
                    rowTemp["EmpName"] = drGetEmpList["EmpName"];
                    rowTemp["IDCardNo"] = drGetEmpList["IDCardNo"];
                    rowTemp["Title"] = drGetEmpList["Title"].ToString().Trim() + "_" + drGetEmpList["ClassTxt"].ToString().Trim();
                    rowTemp["MonthDays"] = (double.TryParse(vCalDays, out vCalDays_D)) ? vCalDays_D.ToString() : "0";
                    rowTemp["BaseAmount"] = (drGetEmpList["Title"].ToString().Trim() == "300") ? vBaseAmount_D : vBaseAmount;
                    rowTemp["Holidays"] = (double.TryParse(vHoliday, out vHolidays_D)) ? vHolidays_D.ToString() : "0";
                    rowTemp["ESCDays"] = (double.TryParse(vESCDays, out vESCDays_D)) ? vESCDays_D.ToString() : "0";
                    rowTemp["WorkDays"] = drGetEmpList["WorkDays"];
                    //vBoundsRatio_D = ((vCalDays_D - vHolidays_D - vESCDays_D < 15.0) || (double.Parse(drGetEmpList["WorkDays"].ToString().Trim()) < 15.0)) ? 0 : 1;
                    vBoundsRatio_D = 1;
                    rowTemp["BoundsRatio"] = vBoundsRatio_D.ToString();
                    vBounds_D = (drGetEmpList["GetPayMode"].ToString().Trim() == "0") ? 0.0 :
                                (drGetEmpList["Title"].ToString().Trim() == "300") ? vBoundsRatio_D * Double.Parse(vBaseAmount_D) :
                                vBoundsRatio_D * Double.Parse(vBaseAmount);
                    rowTemp["Bounds"] = vBounds_D.ToString();
                    rowTemp["WorkType"] = drGetEmpList["WorkType"];
                    rowTemp["Remark"] = drGetEmpList["Remark"];
                    dtTarget.Rows.Add(rowTemp);
                }
            }
        }

        protected void bbCal_Click(object sender, EventArgs e)
        {
            CalData();
            gridResultList.DataSource = dtTarget;
            gridResultList.DataBind();
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            string vCalYM = (Int32.Parse(eCalYM.Text.Trim().Substring(0, 4)) - 1911).ToString() + "年" + Int32.Parse(eCalYM.Text.Trim().Substring(4, 2)).ToString() + "月份";
            CalData();
            ReportDataSource rdsPrint = new ReportDataSource("dsEmpWorkListP", dtTarget);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\EmpWorkListP.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.SetParameters(new ReportParameter("CalYM", vCalYM));
            rvPrint.LocalReport.Refresh();
            rvPrint.Visible = true;
            plReport.Visible = true;
            plCalculater.Visible = false;
            plShow.Visible = false;
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plReport.Visible = false;
            plCalculater.Visible = true;
            plShow.Visible = true;
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            int vLinesNo = 0;
            string vDepNo = "";
            string vDepName = "";
            string vBaseAmount = eCalBase.Text.Trim();
            string vBaseAmount_D = eCalBase_Driver.Text.Trim();
            string vTitle = "";
            string vBounds = "";
            string vFileName = "票價補助發放名冊";
            string vGetPayMode = "";
            string vModeStr = "";
            //string vHeaderStr = "";

            CalData();
            if (dtTarget.Rows.Count > 0)
            {
                //準備匯出EXCEL
                XSSFWorkbook wbExcel = new XSSFWorkbook();
                //Excel 工作表
                XSSFSheet wsExcel = new XSSFSheet();

                //設定標題欄位的格式
                XSSFCellStyle csTitle = (XSSFCellStyle)wbExcel.CreateCellStyle();
                csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csTitle.Alignment = HorizontalAlignment.Center; //水平置中

                //設定標題欄位的字體格式
                XSSFFont fontTitle = (XSSFFont)wbExcel.CreateFont();
                //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
                fontTitle.IsBold = true;
                fontTitle.FontHeightInPoints = 12; //字體大小
                csTitle.SetFont(fontTitle);

                //設定資料內容欄位的格式
                XSSFCellStyle csData = (XSSFCellStyle)wbExcel.CreateCellStyle();
                csData.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                XSSFCellStyle csDoubleData = (XSSFCellStyle)wbExcel.CreateCellStyle();
                csDoubleData.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csDoubleData.Alignment = HorizontalAlignment.Right;

                XSSFDataFormat formatDoubleData = (XSSFDataFormat)wbExcel.CreateDataFormat();
                csDoubleData.DataFormat = formatDoubleData.GetFormat("###,##0.00");

                XSSFFont fontData = (XSSFFont)wbExcel.CreateFont();
                fontData.FontHeightInPoints = 12;
                csData.SetFont(fontData);
                csDoubleData.SetFont(fontData);

                // 2024.07.04 改成不分頁
                /* 2024.07.10 再改回分兩頁，有發跟不發
                vLinesNo = 0;
                //新增一個工作表
                wsExcel = (XSSFSheet)wbExcel.CreateSheet(vFileName);
                //新增標題列
                wsExcel.CreateRow(vLinesNo);
                wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("部門編號");
                wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue("部門");
                wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csTitle;
                wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("工號");
                wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csTitle;
                wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("姓名");
                wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csTitle;
                wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("身分證字號");
                wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csTitle;
                wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("職稱");
                wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csTitle;
                wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("金額");
                wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csTitle;
                wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("備註");
                wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csTitle; //*/

                for (int RowsNo = 0; RowsNo < dtTarget.Rows.Count; RowsNo++)
                {
                    /* 2024.07.04 改成不分頁
                    if (dtTarget.Rows[RowsNo]["DepNo"].ToString().Trim() != vDepNo)
                    {
                        vLinesNo = 0;
                        vDepNo = dtTarget.Rows[RowsNo]["DepNo"].ToString().Trim();
                        vDepName = dtTarget.Rows[RowsNo]["DepName"].ToString().Trim();
                        //新增一個工作表
                        wsExcel = (XSSFSheet)wbExcel.CreateSheet(vDepName);
                        //新增標題列
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("部門編號");
                        wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue("部門");
                        wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("工號");
                        wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("姓名");
                        wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("身分證字號");
                        wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("職稱");
                        wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("金額");
                        wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("備註");
                        wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csTitle;                        
                    } //*/
                    //2024.07.10 再改回分兩頁，有發跟不發
                    if (dtTarget.Rows[RowsNo]["GetPayMode"].ToString().Trim() != vGetPayMode)
                    {
                        vLinesNo = 0;
                        //新增一個工作表
                        vGetPayMode = dtTarget.Rows[RowsNo]["GetPayMode"].ToString().Trim();
                        vModeStr = (vGetPayMode == "1") ? "發放名冊" : "未發放名冊";
                        wsExcel = (XSSFSheet)wbExcel.CreateSheet(vModeStr);
                        
                        //新增標題列
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("部門編號");
                        wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue("部門");
                        wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("工號");
                        wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("姓名");
                        wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("身分證字號");
                        wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("職稱");
                        wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("金額");
                        wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("在職狀況");
                        wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("備註");
                        wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csTitle;
                    }
                    //開始新增資料
                    vLinesNo++;
                    vDepNo = dtTarget.Rows[RowsNo]["DepNo"].ToString().Trim();
                    vDepName = dtTarget.Rows[RowsNo]["DepName"].ToString().Trim();
                    vTitle = dtTarget.Rows[RowsNo]["Title"].ToString().Trim().Substring(0, 3);
                    vBounds = dtTarget.Rows[RowsNo]["Bounds"].ToString().Trim();
                    wsExcel.CreateRow(vLinesNo).CreateCell(0).SetCellValue(vDepNo);
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue(vDepName);
                    wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue(dtTarget.Rows[RowsNo]["EmpNo"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue(dtTarget.Rows[RowsNo]["EmpName"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue(dtTarget.Rows[RowsNo]["IDCardNo"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue(dtTarget.Rows[RowsNo]["Title"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue(vBounds);
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csDoubleData;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue(dtTarget.Rows[RowsNo]["WorkType"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue(dtTarget.Rows[RowsNo]["Remark"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData;
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
                        string vRecordCalYMStr = eCalYM.Text.Trim().Substring(0, 4) + "年" + eCalYM.Text.Trim().Substring(4, 2) + "月";
                        string vRecordNote = "匯出檔案_票價補助發放名冊" + Environment.NewLine +
                                             "EmpWorkList.aspx" + Environment.NewLine +
                                             "計算年月：" + vRecordCalYMStr;
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                        //===========================================================================================================
                        //先判斷是不是IE
                        HttpContext.Current.Response.ContentType = "application/octet-stream";
                        HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                        string TourVerision = brObject.Type;
                        if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                        {
                            HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xlsx", System.Text.Encoding.UTF8));
                        }
                        else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                        {
                            // 設定強制下載標頭
                            Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName + ".xlsx"));
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
                }
            }
        }
    }
}