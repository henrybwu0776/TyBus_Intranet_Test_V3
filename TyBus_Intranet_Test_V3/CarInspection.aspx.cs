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
    public partial class CarInspection : System.Web.UI.Page
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
                    bbCalCheckDate.Visible = (vLoginDepNo == "09");
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        plPrint.Visible = false;
                        eDepNo_S.Enabled = ((vLoginDepNo == "06") || (vLoginDepNo == "09") || (vLoginDepNo == "30") || (vLoginDepNo == "01") || (vLoginDepNo == "02"));
                        eDepNo_E.Enabled = ((vLoginDepNo == "06") || (vLoginDepNo == "09") || (vLoginDepNo == "30") || (vLoginDepNo == "01") || (vLoginDepNo == "02"));
                        eDepNo_S.Text = ((vLoginDepNo == "06") || (vLoginDepNo == "09") || (vLoginDepNo == "30") || (vLoginDepNo == "01") || (vLoginDepNo == "02")) ? "" : vLoginDepNo;
                    }
                    else
                    {
                        DataListBind();
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
        /// 取回 SQL 查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelStr()
        {
            string vResultStr = "";
            string vWStr_Temp = "";
            string vWStr_DepNo = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            switch (rbSearchMode.SelectedValue)
            {
                case "0":
                    vWStr_Temp = "";
                    break;
                case "1":
                    //2019.12.23 原本判斷已過期是以 "下次檢驗日" 做依據，修車廠岳專提出修改為以 "檢驗期限" 來判定是不是過期
                    //vWStr_Temp = "   and c.NexteDate < GetDate() " + Environment.NewLine;
                    vWStr_Temp = " where NCheckTerm < GetDate() " + Environment.NewLine;
                    break;
                case "2":
                    vWStr_Temp = " where DateDiff(day, GetDate(), z.NexteDate) between 0 and 30" + Environment.NewLine;
                    break;
            }
            if (rbSearchMode.SelectedValue == "0")
            {

                vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? " where z.CompanyNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine :
                              ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? " where z.CompanyNo = '" + eDepNo_S.Text.Trim() + "' " + Environment.NewLine :
                              ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? " where z.CompanyNo = '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine : "";
                vResultStr = "select * " + Environment.NewLine +
                             "  from ( " + Environment.NewLine +
                             "        select ca.CompanyNo, (select [Name] from Department where Department.DepNo = ca.CompanyNo) DepName, " + Environment.NewLine +
                             "               ca.Car_ID, ca.Car_No, ca.Point, " + Environment.NewLine +
                             "               (select ClassTxt from DBDICB where ClassNo = ca.Point and FKey = '車輛資料作業    Car_infoA       POINT') Point_C, " + Environment.NewLine +
                             // 2020.07.13 應修車廠岳專要求增加 "可驗車日期" 的欄位
                             //"               ca.ProdDate, ca.NCheckTerm, ca.NextEDate, cc.LastCheckDate, " + Environment.NewLine +
                             // 修正 2022.02.17 修車廠轉站上反應 "可驗車日期" 會提早一天的問題 
                             //"               ca.ProdDate, ca.NCheckTerm, ca.NextEDate, cc.LastCheckDate, DateAdd(day, -29, ca.NextEDate) StartDate, " + Environment.NewLine +
                             "               ca.ProdDate, ca.NCheckTerm, ca.NextEDate, cc.LastCheckDate, " + Environment.NewLine +
                             "               case when DateAdd(day, -29, ca.NextEDate) > DateAdd(month, -1, NextEDate) then DateAdd(day, -29, ca.NextEDate) else DateAdd(month, -1, NextEDate) end StartDate, " + Environment.NewLine +
                             //變更月數計算方式
                             //"               DateDiff(month, ca.NextEDate, cc.LastCheckDate) CheckDateDiff, " + Environment.NewLine +
                             "               Round((DateDiff(day, cc.LastCheckDate, ca.NextEDate) / 30.0), 2) CheckDateDiff, " + Environment.NewLine +
                             "               (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CAR_CLASS' and ClassNo = ca.Car_Class) Car_Class_C, " + Environment.NewLine +
                             "               Car_TypeID, GetLicDate, " + Environment.NewLine +
                             "               (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CLASS' and ClassNo = ca.Class) Class_C, " + Environment.NewLine +
                             "               case when ca.Remark like '%監理站報停%' then '監理站報停' " + Environment.NewLine +
                             "                    else (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       TRAN_TYPE' and ClassNo = ca.Tran_Type) end Tran_Type_C, " + Environment.NewLine +
                             "               ca.Driver, (select[Name] from Employee where EmpNo = ca.Driver) DriverName, " + Environment.NewLine +
                             "               ca.Driver1, (select[Name] from Employee where EmpNo = ca.Driver1) DriverName1, " + Environment.NewLine +
                             "               DateDiff(month, ca.ProdDate, GetDate()) ProdMonth, DateDiff(month, ca.GetLicDate, GetDate()) UsedMonth, " + Environment.NewLine +
                             "               case when ca.Point in ('X', 'X2') then " + Environment.NewLine +
                             "                         case when DateDiff(month, ca.GetLicDate, GetDate()) <= 60 then 60 " + Environment.NewLine +
                             "                              when DateDiff(month, ca.GetLicDate, GetDate()) > 60 and DateDiff(month, ca.GetLicDate, GetDate()) <= 120 then 12 " + Environment.NewLine +
                             "                              else 6 end " + Environment.NewLine +
                             "               else " + Environment.NewLine +
                             "                         case when DateDiff(month, ca.GetLicDate, GetDate()) <= 60 then 12 " + Environment.NewLine +
                             "                              when DateDiff(month, ca.GetLicDate, GetDate()) > 60 and DateDiff(month, ca.GetLicDate, GetDate()) <= 120 then 6 " + Environment.NewLine +
                             "                              else 4 end " + Environment.NewLine +
                             "               end Intervals " + Environment.NewLine +
                             "          from Car_infoA ca left join(select MAX(CheckDate) LastCheckDate, Car_ID from Car_InfoC group by Car_ID) as cc on cc.Car_ID = ca.Car_ID " + Environment.NewLine +
                             "         where ca.Tran_Type in ('1', '2') " + Environment.NewLine +
                             ") z " + Environment.NewLine +
                             vWStr_Temp +
                             vWStr_DepNo +
                             " order by z.CompanyNo, z.Car_ID";
            }
            else
            {

                vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "   and z.CompanyNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine :
                              ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? "   and z.CompanyNo = '" + eDepNo_S.Text.Trim() + "' " + Environment.NewLine :
                              ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? "   and z.CompanyNo = '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine : "";
                vResultStr = "select * " + Environment.NewLine +
                             "  from ( " + Environment.NewLine +
                             "        select ca.CompanyNo, (select [Name] from Department where Department.DepNo = ca.CompanyNo) DepName, " + Environment.NewLine +
                             "               ca.Car_ID, ca.Car_No, ca.Point, " + Environment.NewLine +
                             "               (select ClassTxt from DBDICB where ClassNo = ca.Point and FKey = '車輛資料作業    Car_infoA       POINT') Point_C, " + Environment.NewLine +
                             // 2020.07.13 應修車廠岳專要求增加 "可驗車日期" 的欄位
                             //"               ca.ProdDate, ca.NCheckTerm, ca.NextEDate, cc.LastCheckDate, " + Environment.NewLine +
                             // 修正 2022.02.17 修車廠轉站上反應 "可驗車日期" 會提早一天的問題 
                             //"               ca.ProdDate, ca.NCheckTerm, ca.NextEDate, cc.LastCheckDate, DateAdd(day, -29, ca.NextEDate) StartDate, " + Environment.NewLine +
                             "               ca.ProdDate, ca.NCheckTerm, ca.NextEDate, cc.LastCheckDate, " + Environment.NewLine +
                             "               case when DateAdd(day, -29, ca.NextEDate) > DateAdd(month, -1, NextEDate) then DateAdd(day, -29, ca.NextEDate) else DateAdd(month, -1, NextEDate) end StartDate, " + Environment.NewLine +
                             "               DateDiff(month, ca.NextEDate, cc.LastCheckDate) CheckDateDiff, " + Environment.NewLine +
                             "               (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CAR_CLASS' and ClassNo = ca.Car_Class) Car_Class_C, " + Environment.NewLine +
                             "               Car_TypeID, GetLicDate, " + Environment.NewLine +
                             "               (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CLASS' and ClassNo = ca.Class) Class_C, " + Environment.NewLine +
                             "               (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       TRAN_TYPE' and ClassNo = ca.Tran_Type) Tran_Type_C, " + Environment.NewLine +
                             "               ca.Driver, (select[Name] from Employee where EmpNo = ca.Driver) DriverName, " + Environment.NewLine +
                             "               ca.Driver1, (select[Name] from Employee where EmpNo = ca.Driver1) DriverName1, " + Environment.NewLine +
                             "               DateDiff(month, ca.ProdDate, GetDate()) ProdMonth, DateDiff(month, ca.GetLicDate, GetDate()) UsedMonth, " + Environment.NewLine +
                             "               case when ca.Point in ('X', 'X2') then " + Environment.NewLine +
                             "                         case when DateDiff(month, ca.GetLicDate, GetDate()) <= 60 then 60 " + Environment.NewLine +
                             "                              when DateDiff(month, ca.GetLicDate, GetDate()) > 60 and DateDiff(month, ca.GetLicDate, GetDate()) <= 120 then 12 " + Environment.NewLine +
                             "                              else 6 end " + Environment.NewLine +
                             "               else " + Environment.NewLine +
                             "                         case when DateDiff(month, ca.GetLicDate, GetDate()) <= 60 then 12 " + Environment.NewLine +
                             "                              when DateDiff(month, ca.GetLicDate, GetDate()) > 60 and DateDiff(month, ca.GetLicDate, GetDate()) <= 120 then 6 " + Environment.NewLine +
                             "                              else 4 end " + Environment.NewLine +
                             "               end Intervals " + Environment.NewLine +
                             "          from Car_infoA ca left join(select MAX(CheckDate) LastCheckDate, Car_ID from Car_InfoC group by Car_ID) as cc on cc.Car_ID = ca.Car_ID " + Environment.NewLine +
                             "         where ca.Tran_Type in ('1', '2') " + Environment.NewLine +
                             "           and isnull(DateDiff(month, ca.NextEDate, cc.LastCheckDate), -99) < -1 " + Environment.NewLine +
                             "           and ca.Car_ID not in (select Car_ID from Car_InfoA where Remark like '%監理站報停%') " + Environment.NewLine +
                             ") z " + Environment.NewLine + 
                             vWStr_Temp + 
                             vWStr_DepNo +
                             " order by z.CompanyNo, z.Car_ID";
            }
            return vResultStr;
        }

        /// <summary>
        /// 進行資料繫結
        /// </summary>
        private void DataListBind()
        {
            string vSelectStr = GetSelStr();
            sdsCarInspectionList.SelectCommand = "";
            sdsCarInspectionList.SelectCommand = vSelectStr;
            gridCarInspectionList.DataBind();
        }

        /// <summary>
        /// 取得匯出 EXCEL 檔用的 SQL 查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetExportSelStr()
        {
            string vResultStr = "";
            string vWStr_Temp = "";
            string vWStr_DepNo = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            switch (rbSearchMode.SelectedValue)
            {
                case "0":
                    vWStr_Temp = "";
                    break;
                case "1":
                    //2019.12.23 原本判斷已過期是以 "下次檢驗日" 做依據，修車廠岳專提出修改為以 "檢驗期限" 來判定是不是過期
                    //vWStr_Temp = "   and c.NexteDate < GetDate() " + Environment.NewLine;
                    vWStr_Temp = "   and c.NCheckTerm < GetDate() " + Environment.NewLine;
                    break;
                case "2":
                    vWStr_Temp = "   and DateDiff(day, GetDate(), c.NexteDate) between 0 and 30" + Environment.NewLine;
                    break;
            }
            vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "   and c.CompanyNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine :
                          ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? "   and c.CompanyNo = '" + eDepNo_S.Text.Trim() + "' " + Environment.NewLine :
                          ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? "   and c.CompanyNo = '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine : "";

            if (rbSearchMode.SelectedValue == "0")
            {
                vResultStr = "SELECT c.CompanyNo, (SELECT [NAME] FROM DEPARTMENT WHERE(DEPNO = c.CompanyNo)) AS CompanyName, c.Car_ID, " + Environment.NewLine +
                             //2020.07.13 應修車廠岳專要求，增加 "可驗車日期" 的欄位
                             "       CONVERT(varchar(10), DateAdd(Day, -29, NextEDate), 111) as StartDate, " + Environment.NewLine +
                             "       CONVERT(varchar(10), nextedate, 111) AS NextEDate, " + Environment.NewLine +
                             "       CONVERT(varchar(10), NCheckTerm, 111) AS NCheckTerm, " + Environment.NewLine +
                             "       Remark " + Environment.NewLine +
                             "  FROM Car_infoA c left join(select MAX(CheckDate) LastCheckDate, Car_ID from Car_InfoC group by Car_ID) as cc on cc.Car_ID = c.Car_ID " + Environment.NewLine +
                             " where c.Tran_Type in ('1' ,'2') " + Environment.NewLine + vWStr_Temp + vWStr_DepNo +
                             " order by c.CompanyNo, c.Car_ID";
            }
            else
            {
                vResultStr = "SELECT c.CompanyNo, (SELECT [NAME] FROM DEPARTMENT WHERE(DEPNO = c.CompanyNo)) AS CompanyName, c.Car_ID, " + Environment.NewLine +
                             //2020.07.13 應修車廠岳專要求，增加 "可驗車日期" 的欄位
                             "       CONVERT(varchar(10), DateAdd(Day, -29, NextEDate), 111) as StartDate, " + Environment.NewLine +
                             "       CONVERT(varchar(10), nextedate, 111) AS NextEDate, " + Environment.NewLine +
                             "       CONVERT(varchar(10), NCheckTerm, 111) AS NCheckTerm, " + Environment.NewLine +
                             "       Remark " + Environment.NewLine +
                             "  FROM Car_infoA c left join(select MAX(CheckDate) LastCheckDate, Car_ID from Car_InfoC group by Car_ID) as cc on cc.Car_ID = c.Car_ID " + Environment.NewLine +
                             " where c.Tran_Type in ('1' ,'2') " + Environment.NewLine + vWStr_Temp + vWStr_DepNo +
                             "   and c.Car_ID not in (select Car_ID from Car_InfoA where Remark like '%監理站報停%') " + Environment.NewLine +
                             "   and isnull(DateDiff(month, c.NextEDate, cc.LastCheckDate), -99) < -1 " + Environment.NewLine +
                             " order by c.CompanyNo, c.Car_ID";
            }
            return vResultStr;
        }

        /// <summary>
        /// 重算檢驗日
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCalCheckDate_Click(object sender, EventArgs e)
        {
            string vCarID = "";
            string vTranType = "";
            string vRemark = "";
            string vPoint = "";
            DateTime vGetLicDate;
            DateTime vNextEDate;
            DateTime vNCheckTerm;
            string vNextEDateStr = "";
            string vNCheckTermStr = "";
            int vUsedMonth = 0;
            int vCheckCircle = 0;
            string vRemarkStr = "監理站報停";
            string vUpdateStr = "";

            DateTime vToday = DateTime.Today;
            DateTime vTempDate;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            //======================================================================================================
            string vRecordNote = "計算數值_車輛檢驗到期查核_重新計算檢驗日" + Environment.NewLine +
                                 "CarInspection.aspx";
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            //======================================================================================================
            string vSQLStr = "select Car_ID, Tran_Type, GetLicDate, (DateDiff(month, GetLicDate, GetDate())) UsedMonth, Remark, Point " + Environment.NewLine +
                             "  from Car_InfoA where Tran_Type in ('1', '2') " + Environment.NewLine +
                             " order by Car_ID";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                while (drTemp.Read())
                {
                    vCarID = drTemp["Car_ID"].ToString().Trim();
                    vTranType = drTemp["Tran_Type"].ToString().Trim();
                    vGetLicDate = DateTime.Parse(drTemp["GetLicDate"].ToString().Trim());
                    vUsedMonth = Int32.Parse(drTemp["UsedMonth"].ToString().Trim());
                    vRemark = drTemp["Remark"].ToString().Trim();
                    vPoint = drTemp["Point"].ToString().Trim();
                    if ((vTranType == "1") || ((vTranType == "2") && (vRemark.IndexOf(vRemarkStr) == -1)))
                    {
                        if ((vPoint == "X") || (vPoint == "X2")) //車種是 "X--公務車" 或是 "X2--救援車"
                        {
                            if (vUsedMonth <= 60) //五年內免驗
                            {
                                vNextEDate = vGetLicDate.AddMonths(60);
                            }
                            else if (vUsedMonth <= 108) //五到十年一年一驗
                            {
                                vCheckCircle = (vUsedMonth / 12);
                                vTempDate = vGetLicDate.AddMonths(vCheckCircle * 12);
                                vNextEDate = (vTempDate < vToday) ? vNextEDate = vTempDate.AddMonths(12) : vTempDate;
                            }
                            else if (vUsedMonth <= 120) //五到十年一年一驗 
                            {
                                vNextEDate = vGetLicDate.AddMonths(120);
                            }
                            else //十年以上一年兩驗
                            {
                                vCheckCircle = (vUsedMonth / 6);
                                vTempDate = vGetLicDate.AddMonths(vCheckCircle * 6);
                                vNextEDate = (vTempDate < vToday) ? vNextEDate = vTempDate.AddMonths(6) : vTempDate;
                            }
                        }
                        else //車種不是 "X--公務車" 或 "X2--救援車"
                        {
                            if (vUsedMonth <= 48) //五年內一年一驗
                            {
                                vCheckCircle = (vUsedMonth / 12);
                                vTempDate = vGetLicDate.AddMonths(vCheckCircle * 12);
                                vNextEDate = (vTempDate < vToday) ? vNextEDate = vTempDate.AddMonths(12) : vTempDate;
                            }
                            else if (vUsedMonth <= 60) //五年內一年一驗
                            {
                                vNextEDate = vGetLicDate.AddMonths(60);
                            }
                            else if (vUsedMonth <= 114) //五到十年一年兩驗
                            {
                                vCheckCircle = (vUsedMonth / 6);
                                vTempDate = vGetLicDate.AddMonths(vCheckCircle * 6);
                                vNextEDate = (vTempDate < vToday) ? vNextEDate = vTempDate.AddMonths(6) : vTempDate;
                            }
                            else if (vUsedMonth <= 120) //五到十年一年兩驗
                            {
                                vNextEDate = vGetLicDate.AddMonths(120);
                            }
                            else //十年以上一年三驗
                            {
                                vCheckCircle = (vUsedMonth / 4);
                                vTempDate = vGetLicDate.AddMonths(vCheckCircle * 4);
                                vNextEDate = (vTempDate < vToday) ? vNextEDate = vTempDate.AddMonths(4) : vTempDate;
                            }
                        }
                        vNCheckTerm = vNextEDate.AddDays(-15);
                        vNextEDateStr = vNextEDate.Year.ToString("D4") + "/" + vNextEDate.ToString("MM/dd");
                        vNCheckTermStr = vNCheckTerm.Year.ToString("D4") + "/" + vNCheckTerm.ToString("MM/dd");
                        if (vUpdateStr == "")
                        {
                            vUpdateStr = "update Car_InfoA set NextEDate = '" + vNextEDateStr + "', NCheckTerm = '" + vNCheckTermStr + "' " + Environment.NewLine +
                                         " where Car_ID = '" + vCarID + "' ";
                        }
                        else
                        {
                            vUpdateStr = vUpdateStr + Environment.NewLine +
                                         "update Car_InfoA set NextEDate = '" + vNextEDateStr + "', NCheckTerm = '" + vNCheckTermStr + "' " + Environment.NewLine +
                                         " where Car_ID = '" + vCarID + "' ";
                        }
                    }
                }
                if (vUpdateStr != "")
                {
                    PF.ExecSQL(vConnStr, vUpdateStr);
                }
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            DataListBind();
        }

        /// <summary>
        /// 匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExcel_Click(object sender, EventArgs e)
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
            DateTime vBuDate;

            string vSelectStr = GetExportSelStr();
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
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(rbSearchMode.SelectedItem.Text.Trim());
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i) == "CompanyNo") ? "站別代號" :
                                      (drExcel.GetName(i) == "CompanyName") ? "站別" :
                                      (drExcel.GetName(i) == "Car_ID") ? "牌照號碼" :
                                      (drExcel.GetName(i) == "StartDate") ? "可驗車日期" :
                                      (drExcel.GetName(i) == "NCheckTerm") ? "下次期限" :
                                      (drExcel.GetName(i) == "NextEDate") ? "下次檢驗日" :
                                      (drExcel.GetName(i) == "Remark") ? "備註" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i) == "NCheckTerm") ||
                                 (drExcel.GetName(i) == "NextEDate") ||
                                 (drExcel.GetName(i) == "StartDate")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                if (vBuDate.Year > 3822)
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(((vBuDate.Year) - 3822).ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
                                }
                                else if (vBuDate.Year > 1911)
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(((vBuDate.Year) - 1911).ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
                                }
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
                            string vRecordDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_S.Text.Trim() + "~" + eDepNo_E.Text.Trim() :
                                                     ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() :
                                                     ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() : "全部";
                            string vRecordModeStr = rbSearchMode.SelectedItem.Text.Trim() + "年度";
                            string vRecordNote = "匯出檔案_車輛檢驗到期查核" + Environment.NewLine +
                                                 "CarInspection.aspx" + Environment.NewLine +
                                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                                 "查核類別：" + vRecordModeStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            string vFileName = "車輛檢驗到期查核";
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

        /// <summary>
        /// 輸出報表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrint_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                string vReportName = "";

                if (rbSearchMode.SelectedValue != "0")
                {
                    vReportName = rbSearchMode.SelectedItem.Text.Trim() + "清冊";
                }

                string vRecordDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_S.Text.Trim() + "~" + eDepNo_E.Text.Trim() :
                                         ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() :
                                         ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() : "全部";
                string vRecordModeStr = rbSearchMode.SelectedItem.Text.Trim() + "年度";
                string vRecordNote = "列印報表_車輛檢驗到期查核" + Environment.NewLine +
                                     "CarInspection.aspx" + Environment.NewLine +
                                     "站別：" + vRecordDepNoStr + Environment.NewLine +
                                     "查核類別：" + vRecordModeStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);

                string vSelectStr = GetExportSelStr();
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable();
                daPrint.Fill(dtPrint);

                ReportDataSource rdsPrint = new ReportDataSource("CarInspectionP", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\CarInspection.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plShowData.Visible = false;
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 結束預覽報表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCLoseReport_Click(object sender, EventArgs e)
        {
            plShowData.Visible = true;
            plPrint.Visible = false;
        }
    }
}