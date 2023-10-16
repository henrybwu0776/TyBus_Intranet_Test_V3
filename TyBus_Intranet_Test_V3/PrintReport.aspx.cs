using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;

namespace TyBus_Intranet_Test_V3
{
    public partial class PrintReport : System.Web.UI.Page
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
                    //動態掛載日期輸入視窗
                    string vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eDay_B_Search.ClientID;
                    string vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eDay_B_Search.Attributes["onClick"] = vDateScript;

                    vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eDay_E_Search.ClientID;
                    vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eDay_E_Search.Attributes["onClick"] = vDateScript;

                    if (!IsPostBack)
                    {
                        eDay_B_Search.Text = "";
                        eDay_E_Search.Text = "";
                    }
                    else
                    {
                        BindingPrintData();
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

        private string GetSelStr()
        {
            string vReturnStr = "";
            string vWStr_DepNo = "";
            string vWStr_Day = "";

            //過濾部門別
            vWStr_DepNo = ((eDepNo_B_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() != "")) ? "   AND DepNo between '" + eDepNo_B_Search.Text.Trim() + "' AND '" + eDepNo_E_Search.Text.Trim() + "' " + Environment.NewLine :
                          ((eDepNo_B_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() == "")) ? "   AND DepNo = '" + eDepNo_B_Search.Text.Trim() + "' " + Environment.NewLine :
                          ((eDepNo_B_Search.Text.Trim() == "") && (eDepNo_E_Search.Text.Trim() != "")) ? "   AND DepNo = '" + eDepNo_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            //過濾申請日期
            vWStr_Day = ((eDay_B_Search.Text.Trim() != "") && (eDay_E_Search.Text.Trim() != "")) ? "   AND Day between '" + DateTime.Parse(eDay_B_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eDay_B_Search.Text.Trim()).ToString("MM/dd") + "' AND '" + DateTime.Parse(eDay_E_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eDay_E_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                        ((eDay_B_Search.Text.Trim() != "") && (eDay_E_Search.Text.Trim() == "")) ? "   AND Day = '" + DateTime.Parse(eDay_B_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eDay_B_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                        ((eDay_B_Search.Text.Trim() == "") && (eDay_E_Search.Text.Trim() != "")) ? "   AND Day = '" + DateTime.Parse(eDay_E_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eDay_E_Search.Text.Trim()).ToString("MM/dd") + Environment.NewLine : "";
            switch (rbGroupBy.SelectedIndex)
            {
                case 0: //每張單分開，不合併
                    vReturnStr = "select SheetNo 報修單號, (select Name from Department where Department.DepNo = s.DepNo) 報修單位, " + Environment.NewLine +
                                 "       BuMan, (select Name from Employee where Employee.EmpNo = s.BuMan) as 申請人, " + Environment.NewLine +
                                 "       case when isnull(FixType,'01') = '01' then '軟體' " + Environment.NewLine +
                                 "            when isnull(FixType,'01') = '02' then '硬體' " + Environment.NewLine +
                                 "            else '資訊評估' end 報修類別, " + Environment.NewLine +
                                 "       case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end 報修項目, " + Environment.NewLine +
                                 "       FixRemark 報修內容, convert(varchar, AssignDate, 111) 派工日期, " + Environment.NewLine +
                                 "       (select Name from Employee where Employee.EmpNo = s.Handler) 處理人員, " + Environment.NewLine +
                                 "       case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "            when isnull(Disposal, '01') = '02' then '施工中' " + Environment.NewLine +
                                 "            else '已完工' end 處理進度, convert(varchar, FixFinishDate, 111) 完工日期 " + Environment.NewLine +
                                 "  from SheetA s where SheetType = 'Z' " + vWStr_DepNo + vWStr_Day;
                    break;
                case 1: //依申請人合併
                    vReturnStr = "select count(s.SheetNo) 筆數, (select Name from Employee where Employee.EmpNo = s.BuMan) 申請人, " + Environment.NewLine +
                                 "       FixType 報修類別, FixNote 報修項目, convert(varchar, AssignDate, 111) 派工日期, Handler 處理人員, Disposal 處理進度, convert(varchar, FixFinishDate, 111) 完工日期 " + Environment.NewLine +
                                 "  from (" + Environment.NewLine +
                                 "        select SheetNo, BuMan, " + Environment.NewLine +
                                 "               case when isnull(FixType, '01') = '01' then '軟體' " + Environment.NewLine +
                                 "                    when isnull(FixType, '01') = '02' then '硬體' " + Environment.NewLine +
                                 "                    else '資訊評估' end as FixType, " + Environment.NewLine +
                                 "               case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end as FixNote, " + Environment.NewLine +
                                 "               AssignDate, (select Name from Employee where Employee.EmpNo = SheetA.Handler) as Handler, " + Environment.NewLine +
                                 "               case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "                    when isnull(Disposal, '01') = '02' then '施工中' " + Environment.NewLine +
                                 "                    else '已完工' end as Disposal, FixFinishDate " + Environment.NewLine +
                                 "          from SheetA where SheetType = 'Z'" + Environment.NewLine + vWStr_DepNo + vWStr_Day +
                                 "        ) s " + Environment.NewLine +
                                 " group by BuMan, FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate";
                    break;
                case 2: //依部門合併
                    vReturnStr = "select count(s.SheetNo) 筆數, (select Name from Department where Department.DepNo = s.DepNo) 申請單位, " + Environment.NewLine +
                                 "       FixType 報修類別, FixNote 報修項目, AssignDate 派工日期, Handler 處理人員, Disposal 處理進度, FixFinishDate 完工日期 " + Environment.NewLine +
                                 "  from (" + Environment.NewLine +
                                 "        select SheetNo, DepNo, " + Environment.NewLine +
                                 "               case when isnull(FixType, '01') = '01' then '軟體' " + Environment.NewLine +
                                 "                    when isnull(FixType, '01') = '02' then '硬體' " + Environment.NewLine +
                                 "                    else '資訊評估' end as FixType, " + Environment.NewLine +
                                 "               case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end as FixNote, " + Environment.NewLine +
                                 "               AssignDate, (select Name from Employee where Employee.EmpNo = SheetA.Handler) as Handler, " + Environment.NewLine +
                                 "               case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "                    when isnull(Disposal, '01') = '02' then '施工中' " + Environment.NewLine +
                                 "                    else '已完工' end as Disposal, FixFinishDate " + Environment.NewLine +
                                 "          from SheetA where SheetType = 'Z'" + Environment.NewLine + vWStr_DepNo + vWStr_Day +
                                 "        ) s " + Environment.NewLine +
                                 " group by DepNo, FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate";
                    break;
                case 3: //依日期合併
                    vReturnStr = "select count(s.SheetNo) 筆數, Day 申請日期, FixType 報修類別, FixNote 報修項目, " + Environment.NewLine +
                                 "       convert(varchar, AssignDate, 111) 派工日期, Handler 處理人員, Disposal 處理進度, convert(varchar, FixFinishDate, 111) 完工日期 " + Environment.NewLine +
                                 "  from (" + Environment.NewLine +
                                 "        select SheetNo, convert(varchar, Day, 111) Day, " + Environment.NewLine +
                                 "               case when isnull(FixType,'01') = '01' then '軟體' " + Environment.NewLine +
                                 "                    when isnull(FixType,'01') = '02' then '硬體' " + Environment.NewLine +
                                 "                    else '資訊評估' end as FixType, " + Environment.NewLine +
                                 "               case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end as FixNote, " + Environment.NewLine +
                                 "               AssignDate, (select Name from Employee where Employee.EmpNo = SheetA.Handler) as Handler, " + Environment.NewLine +
                                 "               case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "                    when isnull(Disposal,'01') = '02' then '施工中' " + Environment.NewLine +
                                 "                    else '已完工' end as Disposal, FixFinishDate " + Environment.NewLine +
                                 "          from SheetA where SheetType = 'Z'" + Environment.NewLine + vWStr_DepNo + vWStr_Day +
                                 "        ) s " + Environment.NewLine +
                                 " group by Day, FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate";
                    break;
                case 4: //依申請人及日期合併
                    vReturnStr = "select count(s.SheetNo) 筆數, Day 申請日期, (select Name from Employee where Employee.EmpNo = s.BuMan) 申請人, " + Environment.NewLine +
                                 "       FixType 報修類別, FixNote 報修項目, convert(varchar, AssignDate, 111) 派工日期, Handler 處理人員, " + Environment.NewLine +
                                 "       Disposal 處理進度, convert(varchar, FixFinishDate, 111) 完工日期 " + Environment.NewLine +
                                 "  from (" + Environment.NewLine +
                                 "        select SheetNo, convert(varchar, Day, 111) Day, BuMan, " + Environment.NewLine +
                                 "               case when isnull(FixType, '01') = '01' then '軟體' " + Environment.NewLine +
                                 "                    when isnull(FixType, '01') = '02' then '硬體' " + Environment.NewLine +
                                 "                    else '資訊評估' end as FixType, " + Environment.NewLine +
                                 "               case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end as FixNote, " + Environment.NewLine +
                                 "               AssignDate, (select Name from Employee where Employee.EmpNo = SheetA.Handler) as Handler, " + Environment.NewLine +
                                 "               case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "                    when isnull(Disposal, '01') = '02' then '施工中' " + Environment.NewLine +
                                 "                    else '已完工' end as Disposal, FixFinishDate " + Environment.NewLine +
                                 "          from SheetA where SheetType = 'Z'" + Environment.NewLine + vWStr_DepNo + vWStr_Day +
                                 "        ) s " + Environment.NewLine +
                                 " group by Day, BuMan, FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate";
                    break;
                case 5: //依部門及日期合併
                    vReturnStr = "select count(s.SheetNo) 筆數, Day 申請日期, (select Name from Department where Department.DepNo = s.DepNo) 申請單位, " + Environment.NewLine +
                                 "       FixType 報修類別, FixNote 報修項目, convert(varchar, AssignDate, 111) 派工日期, Handler 處理人員, " + Environment.NewLine +
                                 "       Disposal 處理進度, convert(varchar, FixFinishDate, 111) 完工日期 " + Environment.NewLine +
                                 "  from (" + Environment.NewLine +
                                 "        select SheetNo, convert(varchar, Day, 111) Day, DepNo, " + Environment.NewLine +
                                 "               case when isnull(FixType, '01') = '01' then '軟體' " + Environment.NewLine +
                                 "                    when isnull(FixType, '01') = '02' then '硬體' " + Environment.NewLine +
                                 "                    else '資訊評估' end as FixType, " + Environment.NewLine +
                                 "               case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end as FixNote, " + Environment.NewLine +
                                 "               AssignDate, (select Name from Employee where Employee.EmpNo = SheetA.Handler) as Handler, " + Environment.NewLine +
                                 "               case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "                    when isnull(Disposal, '01') = '02' then '施工中' " + Environment.NewLine +
                                 "                    else '已完工' end as Disposal, FixFinishDate " + Environment.NewLine +
                                 "          from SheetA where SheetType = 'Z' " + Environment.NewLine + vWStr_DepNo + vWStr_Day +
                                 "        ) s " + Environment.NewLine +
                                 " group by Day, DepNo, FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate";
                    break;
                case 6: //依處理人員合併
                    vReturnStr = "select count(s.SheetNo) 筆數, Day 申請日期, (select Name from Department where Department.DepNo = s.DepNo) 申請單位, " + Environment.NewLine +
                                 "       Handler 處理人員, convert(varchar, FixFinishDate, 111) 完工日期 " + Environment.NewLine +
                                 "  from (" + Environment.NewLine +
                                 "        select SheetNo, convert(varchar, Day, 111) Day, DepNo, " + Environment.NewLine +
                                 "               case when isnull(FixType,'01') = '01' then '軟體' " + Environment.NewLine +
                                 "                    when isnull(FixType,'01') = '02' then '硬體' " + Environment.NewLine +
                                 "                    else '資訊評估' end as FixType, " + Environment.NewLine +
                                 "               case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end as FixNote, " + Environment.NewLine +
                                 "               AssignDate, (select Name from Employee where Employee.EmpNo = SheetA.Handler) as Handler, " + Environment.NewLine +
                                 "               case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "                    when isnull(Disposal,'01') = '02' then '施工中' " + Environment.NewLine +
                                 "                    else '已完工' end as Disposal, FixFinishDate " + Environment.NewLine +
                                 "          from SheetA where SheetType = 'Z'" + Environment.NewLine + vWStr_DepNo + vWStr_Day +
                                 "        ) s " + Environment.NewLine +
                                 " group by Day, DepNo, FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate";
                    break;
            }
            return vReturnStr;
        }

        private string GetSelStr_P()
        {
            string vResultStr = "";
            string vWStr_DepNo = "";
            string vWStr_Day = "";

            //過濾部門別
            vWStr_DepNo = ((eDepNo_B_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() != "")) ? "   AND DepNo between '" + eDepNo_B_Search.Text.Trim() + "' AND '" + eDepNo_E_Search.Text.Trim() + "' " + Environment.NewLine :
                          ((eDepNo_B_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() == "")) ? "   AND DepNo = '" + eDepNo_B_Search.Text.Trim() + "' " + Environment.NewLine :
                          ((eDepNo_B_Search.Text.Trim() == "") && (eDepNo_E_Search.Text.Trim() != "")) ? "   AND DepNo = '" + eDepNo_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            //過濾申請日期
            vWStr_Day = ((eDay_B_Search.Text.Trim() != "") && (eDay_E_Search.Text.Trim() != "")) ? "   AND Day between '" + DateTime.Parse(eDay_B_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eDay_B_Search.Text.Trim()).ToString("MM/dd") + "' AND '" + DateTime.Parse(eDay_E_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eDay_E_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                        ((eDay_B_Search.Text.Trim() != "") && (eDay_E_Search.Text.Trim() == "")) ? "   AND Day = '" + DateTime.Parse(eDay_B_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eDay_B_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                        ((eDay_B_Search.Text.Trim() == "") && (eDay_E_Search.Text.Trim() != "")) ? "   AND Day = '" + DateTime.Parse(eDay_E_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eDay_E_Search.Text.Trim()).ToString("MM/dd") + Environment.NewLine : "";
            //產生查詢語法
            switch (rbGroupBy.SelectedIndex)
            {
                case 0: //每張單分開，不合併
                    vResultStr = "select SheetNo, (select Name from Department where Department.DepNo = s.DepNo) DepName, " + Environment.NewLine +
                                 "       BuMan, (select Name from Employee where Employee.EmpNo = s.BuMan) as BuMan_C, " + Environment.NewLine +
                                 "       case when isnull(FixType,'01') = '01' then '軟體' " + Environment.NewLine +
                                 "            when isnull(FixType,'01') = '02' then '硬體' " + Environment.NewLine +
                                 "            else '資訊評估' end FixType, " + Environment.NewLine +
                                 "       case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "       case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end FixNote, " + Environment.NewLine +
                                 "       FixRemark, AssignDate, FixFinishDate, " + Environment.NewLine +
                                 "       (select Name from Employee where Employee.EmpNo = s.Handler) Handler_C, " + Environment.NewLine +
                                 "       case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "            when isnull(Disposal, '01') = '02' then '施工中' " + Environment.NewLine +
                                 "            else '已完工' end Disposal_C " + Environment.NewLine +
                                 "  from SheetA s where SheetType = 'Z' " + vWStr_DepNo + vWStr_Day;
                    break;
                case 1: //依申請人合併
                    vResultStr = "select count(s.SheetNo) RCount, (select [Name] from Employee where Employee.EmpNo = s.BuMan) BuMan_C, " + Environment.NewLine +
                                 "       FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate " + Environment.NewLine +
                                 "  from (" + Environment.NewLine +
                                 "        select SheetNo, BuMan, " + Environment.NewLine +
                                 "               case when isnull(FixType, '01') = '01' then '軟體' " + Environment.NewLine +
                                 "                    when isnull(FixType, '01') = '02' then '硬體' " + Environment.NewLine +
                                 "                    else '資訊評估' end as FixType, " + Environment.NewLine +
                                 "               case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end as FixNote, " + Environment.NewLine +
                                 "               AssignDate, (select Name from Employee where Employee.EmpNo = SheetA.Handler) as Handler, " + Environment.NewLine +
                                 "               case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "                    when isnull(Disposal, '01') = '02' then '施工中' " + Environment.NewLine +
                                 "                    else '已完工' end as Disposal, FixFinishDate " + Environment.NewLine +
                                 "          from SheetA where SheetType = 'Z'" + Environment.NewLine + vWStr_DepNo + vWStr_Day +
                                 "        ) s " + Environment.NewLine +
                                 " group by BuMan, FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate" + Environment.NewLine +
                                 " order by BuMan, FixType, FixNote, AssignDate, FixFinishDate ";
                    break;
                case 2: //依部門合併
                    vResultStr = "select count(s.SheetNo) RCount, (select [Name] from Department where Department.DepNo = s.DepNo) DepName, " + Environment.NewLine +
                                 "       FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate " + Environment.NewLine +
                                 "  from (" + Environment.NewLine +
                                 "        select SheetNo, DepNo, " + Environment.NewLine +
                                 "               case when isnull(FixType, '01') = '01' then '軟體' " + Environment.NewLine +
                                 "                    when isnull(FixType, '01') = '02' then '硬體' " + Environment.NewLine +
                                 "                    else '資訊評估' end as FixType, " + Environment.NewLine +
                                 "               case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end as FixNote, " + Environment.NewLine +
                                 "               AssignDate, (select Name from Employee where Employee.EmpNo = SheetA.Handler) as Handler, " + Environment.NewLine +
                                 "               case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "                    when isnull(Disposal, '01') = '02' then '施工中' " + Environment.NewLine +
                                 "                    else '已完工' end as Disposal, FixFinishDate " + Environment.NewLine +
                                 "          from SheetA where SheetType = 'Z'" + Environment.NewLine + vWStr_DepNo + vWStr_Day +
                                 "        ) s " + Environment.NewLine +
                                 " group by DepNo, FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate" + Environment.NewLine +
                                 " order by DepNo, FixType, FixNote, AssignDate, FixFinishDate ";
                    break;
                case 3: //依日期合併
                    vResultStr = "select count(s.SheetNo) RCount, Day, FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate " + Environment.NewLine +
                                 "  from (" + Environment.NewLine +
                                 "        select SheetNo, convert(varchar, Day, 111) Day, " + Environment.NewLine +
                                 "               case when isnull(FixType,'01') = '01' then '軟體' " + Environment.NewLine +
                                 "                    when isnull(FixType,'01') = '02' then '硬體' " + Environment.NewLine +
                                 "                    else '資訊評估' end as FixType, " + Environment.NewLine +
                                 "               case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end as FixNote, " + Environment.NewLine +
                                 "               AssignDate, (select Name from Employee where Employee.EmpNo = SheetA.Handler) as Handler, " + Environment.NewLine +
                                 "               case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "                    when isnull(Disposal,'01') = '02' then '施工中' " + Environment.NewLine +
                                 "                    else '已完工' end as Disposal, FixFinishDate " + Environment.NewLine +
                                 "          from SheetA where SheetType = 'Z'" + Environment.NewLine + vWStr_DepNo + vWStr_Day +
                                 "        ) s " + Environment.NewLine +
                                 " group by Day, FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate" + Environment.NewLine +
                                 " order by Day, FixType, FixNote, AssignDate, FixFinishDate ";
                    break;
                case 4: //依申請人及日期合併
                    vResultStr = "select count(s.SheetNo) RCount, Day, (select [Name] from Employee where Employee.EmpNo = s.BuMan) BuMan_C, " + Environment.NewLine +
                                 "       FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate " + Environment.NewLine +
                                 "  from (" + Environment.NewLine +
                                 "        select SheetNo, convert(varchar, Day, 111) Day, BuMan, " + Environment.NewLine +
                                 "               case when isnull(FixType, '01') = '01' then '軟體' " + Environment.NewLine +
                                 "                    when isnull(FixType, '01') = '02' then '硬體' " + Environment.NewLine +
                                 "                    else '資訊評估' end as FixType, " + Environment.NewLine +
                                 "               case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end as FixNote, " + Environment.NewLine +
                                 "               AssignDate, (select Name from Employee where Employee.EmpNo = SheetA.Handler) as Handler, " + Environment.NewLine +
                                 "               case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "                    when isnull(Disposal, '01') = '02' then '施工中' " + Environment.NewLine +
                                 "                    else '已完工' end as Disposal, FixFinishDate " + Environment.NewLine +
                                 "          from SheetA where SheetType = 'Z'" + Environment.NewLine + vWStr_DepNo + vWStr_Day +
                                 "        ) s " + Environment.NewLine +
                                 " group by Day, BuMan, FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate" + Environment.NewLine +
                                 " order by Day, BuMan, FixType, FixNote, AssignDate, FixFinishDate ";
                    break;
                case 5: //依部門及日期合併
                    vResultStr = "select count(s.SheetNo) RCount, Day, (select [Name] from Department where Department.DepNo = s.DepNo) DepName, " + Environment.NewLine +
                                 "       FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate " + Environment.NewLine +
                                 "  from (" + Environment.NewLine +
                                 "        select SheetNo, convert(varchar, Day, 111) Day, DepNo, " + Environment.NewLine +
                                 "               case when isnull(FixType, '01') = '01' then '軟體' " + Environment.NewLine +
                                 "                    when isnull(FixType, '01') = '02' then '硬體' " + Environment.NewLine +
                                 "                    else '資訊評估' end as FixType, " + Environment.NewLine +
                                 "               case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end as FixNote, " + Environment.NewLine +
                                 "               AssignDate, (select Name from Employee where Employee.EmpNo = SheetA.Handler) as Handler, " + Environment.NewLine +
                                 "               case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "                    when isnull(Disposal, '01') = '02' then '施工中' " + Environment.NewLine +
                                 "                    else '已完工' end as Disposal, FixFinishDate " + Environment.NewLine +
                                 "          from SheetA where SheetType = 'Z' " + Environment.NewLine + vWStr_DepNo + vWStr_Day +
                                 "        ) s " + Environment.NewLine +
                                 " group by Day, DepNo, FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate" + Environment.NewLine +
                                 " order by Day, DepNo, FixType, FixNote, AssignDate, FixFinishDate ";
                    break;
                case 6: //依處理人員合併
                    vResultStr = "select count(s.SheetNo) RCount, Day, (select [Name] from Department where Department.DepNo = s.DepNo) DepName, " + Environment.NewLine +
                                 "       Handler, FixFinishDate " + Environment.NewLine +
                                 "  from (" + Environment.NewLine +
                                 "        select SheetNo, convert(varchar, Day, 111) Day, DepNo, " + Environment.NewLine +
                                 "               case when isnull(FixType,'01') = '01' then '軟體' " + Environment.NewLine +
                                 "                    when isnull(FixType,'01') = '02' then '硬體' " + Environment.NewLine +
                                 "                    else '資訊評估' end as FixType, " + Environment.NewLine +
                                 "               case when isnull(aboutReport, 'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDataModify, 'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutDesign, 'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutHardDriver, 'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSetting, 'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutSurver, 'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutPurchase, 'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                 "               case when isnull(aboutOthers, 'X') = 'V' then '其他' else '' end as FixNote, " + Environment.NewLine +
                                 "               AssignDate, (select Name from Employee where Employee.EmpNo = SheetA.Handler) as Handler, " + Environment.NewLine +
                                 "               case when isnull(Disposal, '01') = '01' then '未派工' " + Environment.NewLine +
                                 "                    when isnull(Disposal,'01') = '02' then '施工中' " + Environment.NewLine +
                                 "                    else '已完工' end as Disposal, FixFinishDate " + Environment.NewLine +
                                 "          from SheetA where SheetType = 'Z'" + Environment.NewLine + vWStr_DepNo + vWStr_Day +
                                 "        ) s " + Environment.NewLine +
                                 " group by Day, DepNo, FixType, FixNote, AssignDate, Handler, Disposal, FixFinishDate" + Environment.NewLine +
                                 " order by Day, DepNo, Handler, FixType, FixNote, AssignDate, FixFinishDate ";
                    break;
            }
            return vResultStr;
        }

        private void BindingPrintData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelStr();
            sdsFixRequestList.SelectCommand = "";
            sdsFixRequestList.SelectCommand = vSelectStr;
            gridFixRequestList.DataSource = sdsFixRequestList;
            gridFixRequestList.AllowPaging = true;
            gridFixRequestList.PageSize = 5;
            gridFixRequestList.DataBind();
            bbPrint.Visible = (gridFixRequestList.Rows.Count > 0);
        }

        protected void rbReportType_SelectedIndexChanged(object sender, EventArgs e)
        {
            DateTime vToday = DateTime.Today;
            switch (rbReportType.SelectedValue)
            {
                case "00": //年度報表
                    eDay_B_Search.Text = PF.GetYearFirstDay(vToday, "C");
                    eDay_E_Search.Text = PF.GetYearLastDay(vToday, "C");
                    break;
                case "01": //月報表
                    eDay_B_Search.Text = PF.GetMonthFirstDay(vToday, "C");
                    eDay_E_Search.Text = PF.GetMonthLastDay(vToday, "C");
                    break;
            }
        }

        protected void eDepNo_B_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo = eDepNo_B_Search.Text.Trim();
            string vDepName = "";
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
            vDepName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName == "")
            {
                vDepName = vDepNo;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName + "' ";
                vDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_B_Search.Text = vDepNo;
            eDepName_B_Search.Text = vDepName;
        }

        protected void eDepNo_E_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo = eDepNo_E_Search.Text.Trim();
            string vDepName = "";
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
            vDepName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName == "")
            {
                vDepName = vDepNo;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName + "' ";
                vDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_E_Search.Text = vDepNo;
            eDepName_E_Search.Text = vDepName;
        }

        protected void rbGroupBy_SelectedIndexChanged(object sender, EventArgs e)
        {
            BindingPrintData();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            plShowData.Visible = true;
            plPrint.Visible = false;
            BindingPrintData();
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelStr_P();
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
               
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable("FixRequestListP");
                daPrint.Fill(dtPrint);

                ReportDataSource rdsPrint = new ReportDataSource("FixRequestListP", dtPrint);
                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = rbGroupBy.SelectedValue;
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.Refresh();
                plShowData.Visible = false;
                plPrint.Visible = true;
            }
        }

        protected void bbClosePrint_Click(object sender, EventArgs e)
        {
            plShowData.Visible = true;
            plPrint.Visible = false;
        }

        protected void gridFixRequestList_PageIndexChanging(object sender, System.Web.UI.WebControls.GridViewPageEventArgs e)
        {
            gridFixRequestList.PageIndex = e.NewPageIndex;
            BindingPrintData();
        }
    }
}