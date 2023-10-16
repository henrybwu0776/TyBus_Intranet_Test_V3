using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class CustomServiceReportList : System.Web.UI.Page
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
        private DateTime vToday;

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
                    vToday = DateTime.Today;
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }

                    string vCalDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCalDate_Start.ClientID;
                    string vCalDateScript = "window.open('" + vCalDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCalDate_Start.Attributes["onClick"] = vCalDateScript;

                    vCalDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCalDate_End.ClientID;
                    vCalDateScript = "window.open('" + vCalDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCalDate_End.Attributes["onClick"] = vCalDateScript;

                    if (!IsPostBack)
                    {
                        plReport.Visible = false;
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

        private string GetSelectStr_DepNo(string fWStr)
        {
            string vResultStr = "select 單位編號, 單位, 反映事項代號, 反映事項, " + Environment.NewLine +
                                "       case when isnull(客訴事項代號 ,'') = '' then 反映事項代號 else 客訴事項代號 END 客訴事項代號, " + Environment.NewLine +
                                "       case when isnull(客訴事項 ,'') = '' then 反映事項 else 客訴事項 END 客訴事項, " + Environment.NewLine +
                                "       case when(isnull(客訴類別代號, '') = '') and(isnull(客訴事項代號, '') = '') then 反映事項代號 " + Environment.NewLine +
                                "            when(isnull(客訴類別代號, '') = '') and(isnull(客訴事項代號, '') <> '') then 客訴事項代號 " + Environment.NewLine +
                                "            else 客訴類別代號 END 客訴類別代號, " + Environment.NewLine +
                                "       case when(isnull(客訴類別, '') = '') and(isnull(客訴事項, '') = '') then 反映事項 " + Environment.NewLine +
                                "            when(isnull(客訴類別, '') = '') and(isnull(客訴事項, '') <> '') then 客訴事項 " + Environment.NewLine +
                                "            else 客訴類別 END 客訴類別, " + Environment.NewLine +
                                "       路線代號, 駕駛員代號, isnull(駕駛員, '不分駕駛員') 駕駛員, count(駕駛員代號) 筆數 " + Environment.NewLine +
                                "  from ( " + Environment.NewLine +
                                "        SELECT ISNULL(AthorityDepNo, 'ZZ') AS 單位編號, ISNULL(AthorityDepNote, '無單位') AS 單位, " + Environment.NewLine +
                                "               ISNULL(ServiceType, '') AS 反映事項代號, (SELECT TypeText FROM CustomServiceType WHERE(TypeStep = '1') AND(TypeLevel1 = a.ServiceType)) AS 反映事項, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeB, '') AS 客訴事項代號, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE(TypeStep = '2') AND(TypeLevel1 = a.ServiceType) AND(TypeLevel2 = a.ServiceTypeB)) AS 客訴事項, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeC, '') AS 客訴類別代號, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE(TypeStep = '3') AND(TypeLevel1 = a.ServiceType) AND(TypeLevel2 = a.ServiceTypeB) AND(TypeLevel3 = a.ServiceTypeC)) AS 客訴類別, " + Environment.NewLine +
                                "               ISNULL(LinesNo, '無路線別') AS 路線代號, " + Environment.NewLine +
                                "               ISNULL(Driver, '000000') AS 駕駛員代號, (SELECT NAME FROM EMPLOYEE WHERE(EMPNO = a.Driver)) AS 駕駛員 " + Environment.NewLine +
                                "          FROM CustomService AS a " + Environment.NewLine +
                                "         where 1 = 1 " + Environment.NewLine + fWStr +
                                "         union all " + Environment.NewLine +
                                "        SELECT ISNULL(AthorityDepNo2, 'ZZ') AS 單位編號, ISNULL(AthorityDepNote2, '無單位') AS 單位, " + Environment.NewLine +
                                "               ISNULL(ServiceType, '') AS 反映事項代號, (SELECT TypeText FROM CustomServiceType WHERE(TypeStep = '1') AND(TypeLevel1 = a.ServiceType)) AS 反映事項, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeB, '') AS 客訴事項代號, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE(TypeStep = '2') AND(TypeLevel1 = a.ServiceType) AND(TypeLevel2 = a.ServiceTypeB)) AS 客訴事項, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeC, '') AS 客訴類別代號, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE(TypeStep = '3') AND(TypeLevel1 = a.ServiceType) AND(TypeLevel2 = a.ServiceTypeB) AND(TypeLevel3 = a.ServiceTypeC)) AS 客訴類別, " + Environment.NewLine +
                                "               ISNULL(LinesNo2, '無路線別') AS 路線代號, " + Environment.NewLine +
                                "               ISNULL(Driver2, '000000') AS 駕駛員代號, (SELECT NAME FROM EMPLOYEE WHERE(EMPNO = a.Driver2)) AS 駕駛員 " + Environment.NewLine +
                                "          FROM CustomService AS a " + Environment.NewLine +
                                "         where isnull(AthorityDepNo2, '') <> '' " + Environment.NewLine + fWStr +
                                ") f " + Environment.NewLine +
                                " group by f.單位編號, f.單位, f.反映事項代號, f.反映事項, f.客訴事項代號, f.客訴事項, f.客訴類別代號, f.客訴類別, f.路線代號, f.駕駛員代號, f.駕駛員 " + Environment.NewLine +
                                " order by f.單位編號, f.反映事項代號, f.客訴事項代號, f.客訴類別代號, f.路線代號, f.駕駛員代號";
            return vResultStr;
        }

        /// <summary>
        /// 依單位別預覽統計報表
        /// </summary>
        /// <param name="fWStr"></param>
        private void PreviewByDepNo(string fWStr)
        {
            string vSelectStr = GetSelectStr_DepNo(fWStr);
            sdsCustomServiceP.SelectCommand = "";
            sdsCustomServiceP.SelectCommand = vSelectStr;
            ReportDataSource rdsPrint = new ReportDataSource("CustomServiceP", sdsCustomServiceP);

            CountDepNo(fWStr);
            ReportDataSource rdsCountDepNo = new ReportDataSource("CountDepNo", sdsCountDepNo);

            CountServiceType(fWStr);
            ReportDataSource rdsCountServiceType = new ReportDataSource("CountServiceType", sdsCountServiceType);

            CountServiceTypeB(fWStr);
            ReportDataSource rdsCountServiceTypeB = new ReportDataSource("CountServiceTypeB", sdsCountServiceTypeB);

            CountServiceTypeC(fWStr);
            ReportDataSource rdsCountServiceTypeC = new ReportDataSource("CountServiceTypeC", sdsCountServiceTypeC);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\ServiceCountByDepNo.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.DataSources.Add(rdsCountDepNo);
            rvPrint.LocalReport.DataSources.Add(rdsCountServiceType);
            rvPrint.LocalReport.DataSources.Add(rdsCountServiceTypeB);
            rvPrint.LocalReport.DataSources.Add(rdsCountServiceTypeC);
            rvPrint.LocalReport.Refresh();
            plReport.Visible = true;
            plShowData.Visible = false;
        }

        private string GetSelectStr_Driver(string fWStr)
        {
            string vResultStr = "SELECT Driver as 駕駛員代號, ISNULL((select [Name] from EMployee where EMpNo = t.Driver and LeaveDay is null), '不分駕駛員') 駕駛員, " + Environment.NewLine +
                                "       AthorityDepNo as 單位編號, AthorityDepNo_C as 單位, " + Environment.NewLine +
                                "       ServiceType as 反映事項代號, ServiceType_C as 反映事項, " + Environment.NewLine +
                                "       case when isnull(ServiceTypeB ,'') = '' then ServiceType else ServiceTypeB END as 客訴事項代號, " + Environment.NewLine +
                                "       case when isnull(ServiceTypeB_C ,'') = '' then ServiceType_C else ServiceTypeB_C END as 客訴事項, " + Environment.NewLine +
                                "       case when(isnull(ServiceTypeC, '') = '') and(isnull(ServiceTypeB, '') = '') then ServiceType " + Environment.NewLine +
                                "            when(isnull(ServiceTypeC, '') = '') and(isnull(ServiceTypeB, '') <> '') then ServiceTypeB " + Environment.NewLine +
                                "            else ServiceTypeC END as 客訴類別代號, " + Environment.NewLine +
                                "       case when(isnull(ServiceTypeC_C, '') = '') and(isnull(ServiceTypeB_C, '') = '') then ServiceType_C " + Environment.NewLine +
                                "            when(isnull(ServiceTypeC_C, '') = '') and(isnull(ServiceTypeB_C, '') <> '') then ServiceTypeB_C " + Environment.NewLine +
                                "            else ServiceTypeC_C END as 客訴類別, " + Environment.NewLine +
                                "       isnull(LinesNo, '不分路線') as 路線代號, " + Environment.NewLine +
                                "       count(LinesNo) as 筆數 " + Environment.NewLine +
                                "  FROM ( " + Environment.NewLine +
                                "        SELECT ISNULL(LinesNo, '無路線別') AS LinesNo, " + Environment.NewLine +
                                "               ISNULL(AthorityDepNo, '') AS AthorityDepNo, isnull((select[Name] from Department where DepNo = a.AthorityDepNo), '不分單位') as AthorityDepNo_C, " + Environment.NewLine +
                                "               ISNULL(ServiceType, '') AS ServiceType, (select TypeText from CustomServiceType where TypeStep = '1' and TypeLevel1 = a.ServiceType) as ServiceType_C, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeB, '') AS ServiceTypeB, (select TypeText from CustomServiceType where TypeStep = '2' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB) as ServiceTypeB_C, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeC, '') AS ServiceTypeC, (select TypeText from CustomServiceType where TypeStep = '3' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB and TypeLevel3 = a.ServiceTypeC) as ServiceTypeC_C, " + Environment.NewLine +
                                "               ISNULL(Driver, '') AS Driver " + Environment.NewLine +
                                "          FROM CustomService a " + Environment.NewLine +
                                "         WHERE 1 = 1 " + Environment.NewLine + fWStr +
                                "         UNION ALL " + Environment.NewLine +
                                "        SELECT ISNULL(LinesNo2, '無路線別') AS LinesNo, " + Environment.NewLine +
                                "               ISNULL(AthorityDepNo2, '') AS AthorityDepNo, isnull((select[Name] from Department where DepNo = a.AthorityDepNo2), '不分單位') as AthorityDepNo_C, " + Environment.NewLine +
                                "               ISNULL(ServiceType, '') AS ServiceType, (select TypeText from CustomServiceType where TypeStep = '1' and TypeLevel1 = a.ServiceType) as ServiceType_C, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeB, '') AS ServiceTypeB, (select TypeText from CustomServiceType where TypeStep = '2' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB) as ServiceTypeB_C, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeC, '') AS ServiceTypeC, (select TypeText from CustomServiceType where TypeStep = '3' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB and TypeLevel3 = a.ServiceTypeC) as ServiceTypeC_C, " + Environment.NewLine +
                                "               ISNULL(Driver2, '') AS Driver " + Environment.NewLine +
                                "          FROM CustomService a " + Environment.NewLine +
                                "         WHERE ISNULL(Driver2, '') <> '' " + Environment.NewLine + fWStr +
                                "       ) t " + Environment.NewLine +
                                " group by t.Driver, t.AthorityDepNo,t.AthorityDepNo_C, t.ServiceType, t.ServiceType_C, t.ServiceTypeB, t.ServiceTypeB_C, t.ServiceTypeC, t.ServiceTypeC_C, t.LinesNo " + Environment.NewLine +
                                " order by t.Driver, t.AthorityDepNo, t.ServiceType, t.ServiceTypeB, t.ServiceTypeC, t.LinesNo";
            return vResultStr;
        }

        /// <summary>
        /// 依駕駛員預覽報表
        /// </summary>
        /// <param name="fWStr"></param>
        private void PreviewByDriver(string fWStr)
        {
            string vSelectStr = GetSelectStr_Driver(fWStr);
            sdsCustomServiceP.SelectCommand = "";
            sdsCustomServiceP.SelectCommand = vSelectStr;
            ReportDataSource rdsPrint = new ReportDataSource("CustomServiceP", sdsCustomServiceP);

            CountDriver(fWStr);
            ReportDataSource rdsCountDriver = new ReportDataSource("CountDriver", sdsCountDriver);

            CountServiceType(fWStr);
            ReportDataSource rdsCountServiceType = new ReportDataSource("CountServiceType", sdsCountServiceType);

            CountServiceTypeB(fWStr);
            ReportDataSource rdsCountServiceTypeB = new ReportDataSource("CountServiceTypeB", sdsCountServiceTypeB);

            CountServiceTypeC(fWStr);
            ReportDataSource rdsCountServiceTypeC = new ReportDataSource("CountServiceTypeC", sdsCountServiceTypeC);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\ServiceCountByDriver.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.DataSources.Add(rdsCountDriver);
            rvPrint.LocalReport.DataSources.Add(rdsCountServiceType);
            rvPrint.LocalReport.DataSources.Add(rdsCountServiceTypeB);
            rvPrint.LocalReport.DataSources.Add(rdsCountServiceTypeC);
            rvPrint.LocalReport.Refresh();
            plReport.Visible = true;
            plShowData.Visible = false;
        }

        /// <summary>
        /// 依類別統計查詢語法
        /// </summary>
        /// <param name="fWStr"></param>
        /// <returns></returns>
        private string GetSelectStr_Type(string fWStr)
        {
            string vResultStr = "SELECT ServiceType as 反映事項代號, ServiceType_C as 反映事項, " + Environment.NewLine +
                                "       case when isnull(ServiceTypeB ,'') = '' then ServiceType else ServiceTypeB END as 客訴事項代號, " + Environment.NewLine +
                                "       case when isnull(ServiceTypeB_C ,'') = '' then ServiceType_C else ServiceTypeB_C END as 客訴事項, " + Environment.NewLine +
                                "       case when(isnull(ServiceTypeC, '') = '') and(isnull(ServiceTypeB, '') = '') then ServiceType " + Environment.NewLine +
                                "            when(isnull(ServiceTypeC, '') = '') and(isnull(ServiceTypeB, '') <> '') then ServiceTypeB " + Environment.NewLine +
                                "            else ServiceTypeC END as 客訴類別代號, " + Environment.NewLine +
                                "       case when(isnull(ServiceTypeC_C, '') = '') and(isnull(ServiceTypeB_C, '') = '') then ServiceType_C " + Environment.NewLine +
                                "            when(isnull(ServiceTypeC_C, '') = '') and(isnull(ServiceTypeB_C, '') <> '') then ServiceTypeB_C " + Environment.NewLine +
                                "            else ServiceTypeC_C END as 客訴類別, " + Environment.NewLine +
                                "       AthorityDepNo as 單位編號, AthorityDepNo_C as 單位, " + Environment.NewLine +
                                "       isnull(LinesNo, '不分路線') as 路線代號, " + Environment.NewLine +
                                "       Driver as 駕駛員代號, ISNULL((select [Name] from EMployee where EMpNo = t.Driver and LeaveDay is null), '不分駕駛員') 駕駛員, " + Environment.NewLine +
                                "       count(Driver) as 筆數 " + Environment.NewLine +
                                "  FROM ( " + Environment.NewLine +
                                "        SELECT ISNULL(LinesNo, '無路線別') AS LinesNo, " + Environment.NewLine +
                                "               ISNULL(AthorityDepNo, '') AS AthorityDepNo, isnull((select[Name] from Department where DepNo = a.AthorityDepNo), '不分單位') as AthorityDepNo_C, " + Environment.NewLine +
                                "               ISNULL(ServiceType, '') AS ServiceType, (select TypeText from CustomServiceType where TypeStep = '1' and TypeLevel1 = a.ServiceType) as ServiceType_C, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeB, '') AS ServiceTypeB, (select TypeText from CustomServiceType where TypeStep = '2' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB) as ServiceTypeB_C, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeC, '') AS ServiceTypeC, (select TypeText from CustomServiceType where TypeStep = '3' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB and TypeLevel3 = a.ServiceTypeC) as ServiceTypeC_C, " + Environment.NewLine +
                                "               ISNULL(Driver, '') AS Driver " + Environment.NewLine +
                                "          FROM CustomService a " + Environment.NewLine +
                                "         WHERE 1 = 1 " + Environment.NewLine + fWStr +
                                "         UNION ALL " + Environment.NewLine +
                                "        SELECT ISNULL(LinesNo2, '無路線別') AS LinesNo, " + Environment.NewLine +
                                "               ISNULL(AthorityDepNo2, '') AS AthorityDepNo, isnull((select[Name] from Department where DepNo = a.AthorityDepNo2), '不分單位') as AthorityDepNo_C, " + Environment.NewLine +
                                "               ISNULL(ServiceType, '') AS ServiceType, (select TypeText from CustomServiceType where TypeStep = '1' and TypeLevel1 = a.ServiceType) as ServiceType_C, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeB, '') AS ServiceTypeB, (select TypeText from CustomServiceType where TypeStep = '2' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB) as ServiceTypeB_C, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeC, '') AS ServiceTypeC, (select TypeText from CustomServiceType where TypeStep = '3' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB and TypeLevel3 = a.ServiceTypeC) as ServiceTypeC_C, " + Environment.NewLine +
                                "               ISNULL(Driver2, '') AS Driver " + Environment.NewLine +
                                "          FROM CustomService a " + Environment.NewLine +
                                "         WHERE ISNULL(Driver2, '') <> '' " + Environment.NewLine + fWStr +
                                "       ) t " + Environment.NewLine +
                                " group by t.ServiceType, t.ServiceType_C, t.ServiceTypeB, t.ServiceTypeB_C, t.ServiceTypeC, t.ServiceTypeC_C, t.AthorityDepNo,t.AthorityDepNo_C, t.LinesNo, t.Driver " + Environment.NewLine +
                                " order by t.ServiceType, t.ServiceTypeB, t.ServiceTypeC, t.AthorityDepNo, t.LinesNo, t.Driver";
            return vResultStr;
        }

        /// <summary>
        /// 依類別預覽報表
        /// </summary>
        /// <param name="fWStr"></param>
        private void PreviewByType(string fWStr)
        {
            string vSelectStr = GetSelectStr_Type(fWStr);
            sdsCustomServiceP.SelectCommand = "";
            sdsCustomServiceP.SelectCommand = vSelectStr;
            ReportDataSource rdsPrint = new ReportDataSource("CustomServiceP", sdsCustomServiceP);

            CountServiceType(fWStr);
            ReportDataSource rdsCountServiceType = new ReportDataSource("CountServiceType", sdsCountServiceType);

            CountServiceTypeB(fWStr);
            ReportDataSource rdsCountServiceTypeB = new ReportDataSource("CountServiceTypeB", sdsCountServiceTypeB);

            CountServiceTypeC(fWStr);
            ReportDataSource rdsCountServiceTypeC = new ReportDataSource("CountServiceTypeC", sdsCountServiceTypeC);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\ServiceCountByServiceType.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.DataSources.Add(rdsCountServiceType);
            rvPrint.LocalReport.DataSources.Add(rdsCountServiceTypeB);
            rvPrint.LocalReport.DataSources.Add(rdsCountServiceTypeC);
            rvPrint.LocalReport.Refresh();
            plReport.Visible = true;
            plShowData.Visible = false;
        }

        /// <summary>
        /// 依路線統計查詢語法
        /// </summary>
        /// <param name="fWStr"></param>
        /// <returns></returns>
        private string GetSelectStr_LinesNo(string fWStr)
        {
            string vReturnStr = "SELECT isnull(LinesNo, '不分路線') as 路線代號, " + Environment.NewLine +
                                "       AthorityDepNo as 單位編號, AthorityDepNo_C as 單位, " + Environment.NewLine +
                                "       ServiceType as 反映事項代號, ServiceType_C as 反映事項, " + Environment.NewLine +
                                "       case when isnull(ServiceTypeB ,'') = '' then ServiceType else ServiceTypeB END as 客訴事項代號, " + Environment.NewLine +
                                "       case when isnull(ServiceTypeB_C ,'') = '' then ServiceType_C else ServiceTypeB_C END as 客訴事項, " + Environment.NewLine +
                                "       case when(isnull(ServiceTypeC, '') = '') and(isnull(ServiceTypeB, '') = '') then ServiceType " + Environment.NewLine +
                                "            when(isnull(ServiceTypeC, '') = '') and(isnull(ServiceTypeB, '') <> '') then ServiceTypeB " + Environment.NewLine +
                                "            else ServiceTypeC END as 客訴類別代號, " + Environment.NewLine +
                                "       case when(isnull(ServiceTypeC_C, '') = '') and(isnull(ServiceTypeB_C, '') = '') then ServiceType_C " + Environment.NewLine +
                                "            when(isnull(ServiceTypeC_C, '') = '') and(isnull(ServiceTypeB_C, '') <> '') then ServiceTypeB_C " + Environment.NewLine +
                                "            else ServiceTypeC_C END as 客訴類別, " + Environment.NewLine +
                                "       Driver as 駕駛員代號, ISNULL((select[Name] from EMployee where EMpNo = t.Driver and LeaveDay is null), '不分駕駛員') 駕駛員, " + Environment.NewLine +
                                "       count(Driver) as 筆數 " + Environment.NewLine +
                                "  FROM ( " + Environment.NewLine +
                                "        SELECT ISNULL(LinesNo, '無路線別') AS LinesNo, " + Environment.NewLine +
                                "               ISNULL(AthorityDepNo, '') AS AthorityDepNo, isnull((select[Name] from Department where DepNo = a.AthorityDepNo), '不分單位') as AthorityDepNo_C, " + Environment.NewLine +
                                "               ISNULL(ServiceType, '') AS ServiceType, (select TypeText from CustomServiceType where TypeStep = '1' and TypeLevel1 = a.ServiceType) as ServiceType_C, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeB, '') AS ServiceTypeB, (select TypeText from CustomServiceType where TypeStep = '2' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB) as ServiceTypeB_C, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeC, '') AS ServiceTypeC, (select TypeText from CustomServiceType where TypeStep = '3' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB and TypeLevel3 = a.ServiceTypeC) as ServiceTypeC_C, " + Environment.NewLine +
                                "               ISNULL(Driver, '') AS Driver " + Environment.NewLine +
                                "          FROM CustomService a " + Environment.NewLine +
                                "         WHERE 1 = 1 " + Environment.NewLine + fWStr +
                                "         UNION ALL " + Environment.NewLine +
                                "        SELECT ISNULL(LinesNo2, '無路線別') AS LinesNo, " + Environment.NewLine +
                                "               ISNULL(AthorityDepNo2, '') AS AthorityDepNo, isnull((select[Name] from Department where DepNo = a.AthorityDepNo2), '不分單位') as AthorityDepNo_C, " + Environment.NewLine +
                                "               ISNULL(ServiceType, '') AS ServiceType, (select TypeText from CustomServiceType where TypeStep = '1' and TypeLevel1 = a.ServiceType) as ServiceType_C, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeB, '') AS ServiceTypeB, (select TypeText from CustomServiceType where TypeStep = '2' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB) as ServiceTypeB_C, " + Environment.NewLine +
                                "               ISNULL(ServiceTypeC, '') AS ServiceTypeC, (select TypeText from CustomServiceType where TypeStep = '3' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB and TypeLevel3 = a.ServiceTypeC) as ServiceTypeC_C, " + Environment.NewLine +
                                "               ISNULL(Driver2, '') AS Driver " + Environment.NewLine +
                                "          FROM CustomService a " + Environment.NewLine +
                                "         WHERE ISNULL(LinesNo2, '') <> '' " + Environment.NewLine + fWStr +
                                "       ) t " + Environment.NewLine +
                                " group by t.LinesNo, t.AthorityDepNo,t.AthorityDepNo_C, t.ServiceType, t.ServiceType_C, t.ServiceTypeB, t.ServiceTypeB_C, t.ServiceTypeC, t.ServiceTypeC_C, t.Driver " + Environment.NewLine +
                                " order by t.LinesNo, t.AthorityDepNo, t.ServiceType, t.ServiceTypeB, t.ServiceTypeC, t.Driver ";
            return vReturnStr;
        }

        /// <summary>
        /// 依路線統計報表
        /// </summary>
        /// <param name="fWStr"></param>
        private void PreviewByLinesNo(string fWStr)
        {
            string vSelectStr = GetSelectStr_LinesNo(fWStr);
            sdsCustomServiceP.SelectCommand = "";
            sdsCustomServiceP.SelectCommand = vSelectStr;
            ReportDataSource rdsPrint = new ReportDataSource("CustomServiceP", sdsCustomServiceP);

            CountLinesNo(fWStr);
            ReportDataSource rdsCountLinesNo = new ReportDataSource("CountLinesNo", sdsCountLinesNo);

            CountServiceType(fWStr);
            ReportDataSource rdsCountServiceType = new ReportDataSource("CountServiceType", sdsCountServiceType);

            CountServiceTypeB(fWStr);
            ReportDataSource rdsCountServiceTypeB = new ReportDataSource("CountServiceTypeB", sdsCountServiceTypeB);

            CountServiceTypeC(fWStr);
            ReportDataSource rdsCountServiceTypeC = new ReportDataSource("CountServiceTypeC", sdsCountServiceTypeC);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\ServiceCountByLinesNo.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.DataSources.Add(rdsCountLinesNo);
            rvPrint.LocalReport.DataSources.Add(rdsCountServiceType);
            rvPrint.LocalReport.DataSources.Add(rdsCountServiceTypeB);
            rvPrint.LocalReport.DataSources.Add(rdsCountServiceTypeC);
            rvPrint.LocalReport.Refresh();
            plReport.Visible = true;
            plShowData.Visible = false;
        }

        /// <summary>
        /// 單位別圖表用
        /// </summary>
        /// <param name="fWStr"></param>
        private void CountDepNo(string fWStr)
        {
            string vSelectStr = "";
            vSelectStr = "SELECT AthorityDepNo, AthorityDepNo_C, COUNT(AthorityDepNo) AS RCount " + Environment.NewLine +
                         "  FROM ( " + Environment.NewLine +
                         "        SELECT ISNULL(AthorityDepNo, 'ZZ') AS AthorityDepNo, " + Environment.NewLine +
                         "               CASE WHEN AthorityDepNo = '99' THEN '其他' ELSE ISNULL(AthorityDepNote, '不分站別') END AS AthorityDepNo_C " + Environment.NewLine +
                         "          FROM CustomService AS a " + Environment.NewLine +
                         "         WHERE 1 = 1 " + Environment.NewLine + fWStr +
                         "         UNION ALL " + Environment.NewLine +
                         "        SELECT ISNULL(AthorityDepNo2, 'ZZ') AS AthorityDepNo, " + Environment.NewLine +
                         "               CASE WHEN AthorityDepNo2 = '99' THEN '其他' ELSE ISNULL(AthorityDepNote2, '不分站別') END AS AthorityDepNo_C " + Environment.NewLine +
                         "          FROM CustomService AS a " + Environment.NewLine +
                         "         WHERE (ISNULL(AthorityDepNo2, N'') <> '') " + Environment.NewLine + fWStr +
                         "       ) AS z " + Environment.NewLine +
                         " GROUP BY   AthorityDepNo, AthorityDepNo_C " + Environment.NewLine +
                         " ORDER BY AthorityDepNo ";
            sdsCountDepNo.SelectCommand = "";
            sdsCountDepNo.SelectCommand = vSelectStr;
        }

        /// <summary>
        /// 駕駛別圖表用
        /// </summary>
        /// <param name="fWStr"></param>
        private void CountDriver(string fWStr)
        {
            string vSelectStr = "";
            vSelectStr = "SELECT Driver, Driver_C, COUNT(Driver) AS RCount " + Environment.NewLine +
                         " FROM ( " + Environment.NewLine +
                         "       SELECT ISNULL(Driver, '000000') AS Driver, " + Environment.NewLine +
                         "              ISNULL((SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = a.Driver)), '不分駕駛員') AS Driver_C " + Environment.NewLine +
                         "         FROM CustomService AS a " + Environment.NewLine +
                         "        WHERE 1 = 1 " + Environment.NewLine + fWStr +
                         "        UNION ALL " + Environment.NewLine +
                         "       SELECT ISNULL(Driver2, '000000') AS Driver, " + Environment.NewLine +
                         "              ISNULL((SELECT [NAME] FROM EMPLOYEE WHERE(EMPNO = a.Driver2)), '不分駕駛員') AS Driver_C " + Environment.NewLine +
                         "         FROM CustomService AS a " + Environment.NewLine +
                         "        WHERE (ISNULL(Driver2, N'') <> '') " + Environment.NewLine + fWStr +
                         "     ) AS z " + Environment.NewLine +
                         " GROUP BY   Driver, Driver_C " + Environment.NewLine +
                         " ORDER BY Driver";
            sdsCountDriver.SelectCommand = "";
            sdsCountDriver.SelectCommand = vSelectStr;
        }

        /// <summary>
        /// 路線別圖表用
        /// </summary>
        /// <param name="fWStr"></param>
        private void CountLinesNo(string fWStr)
        {
            string vSelectStr = "";
            vSelectStr = "SELECT LinesNo, COUNT(LinesNo) AS RCount " + Environment.NewLine +
                         "  FROM ( " + Environment.NewLine +
                         "        SELECT ISNULL(LinesNo, '不分路線') AS LinesNo " + Environment.NewLine +
                         "          FROM CustomService AS a " + Environment.NewLine +
                         "         WHERE (1 = 1) " + Environment.NewLine + fWStr +
                         "         UNION ALL " + Environment.NewLine +
                         "        SELECT LinesNo2 AS LinesNo " + Environment.NewLine +
                         "          FROM CustomService AS a " + Environment.NewLine +
                         "         WHERE (ISNULL(LinesNo2, N'') <> '') " + Environment.NewLine + fWStr +
                         "       ) AS z " + Environment.NewLine +
                         " GROUP BY   LinesNo " + Environment.NewLine +
                         " ORDER BY LinesNo";
        }

        /// <summary>
        /// 反映事項圖表用
        /// </summary>
        /// <param name="fWStr"></param>
        private void CountServiceType(string fWStr)
        {
            string vSelectStr = "";
            vSelectStr = "SELECT ServiceType, " + Environment.NewLine +
                         "       (SELECT TypeText FROM CustomServiceType WHERE(TypeStep = '1') AND(TypeLevel1 = a.ServiceType)) AS ServiceType_C, " + Environment.NewLine +
                         "       COUNT(ServiceNo) AS RCount " + Environment.NewLine +
                         "  FROM CustomService a " + Environment.NewLine +
                         " WHERE 1 = 1 " + Environment.NewLine + fWStr +
                         " GROUP BY ServiceType " + Environment.NewLine +
                         " ORDER BY ServiceType ";
            sdsCountServiceType.SelectCommand = "";
            sdsCountServiceType.SelectCommand = vSelectStr;
        }

        /// <summary>
        /// 客訴事項圖表用
        /// </summary>
        /// <param name="fWStr"></param>
        private void CountServiceTypeB(string fWStr)
        {
            string vSelectStr = "";
            vSelectStr = "SELECT ServiceType, ServiceType_C, COUNT(ServiceNo) AS RCount " + Environment.NewLine +
                         "  FROM ( " + Environment.NewLine +
                         "        SELECT ServiceNo, ServiceType + ISNULL(ServiceTypeB, '000') AS ServiceType, " + Environment.NewLine +
                         "          CASE WHEN ISNULL(ServiceTypeB, '') <> '' THEN " + Environment.NewLine +
                         "                    (SELECT TypeText FROM CustomServiceType WHERE TypeStep = '2' AND TypeLevel1 = a.ServiceType AND TypeLevel2 = a.ServiceTypeB) " + Environment.NewLine +
                         "               ELSE " + Environment.NewLine +
                         "                    (SELECT TypeText FROM CustomserviceType WHERE TypeStep = '1' AND TypeLevel1 = a.ServiceType) END AS ServiceType_C " + Environment.NewLine +
                         "          FROM CustomService AS a " + Environment.NewLine +
                         "         WHERE 1 = 1 " + Environment.NewLine + fWStr +
                         "       ) AS z " + Environment.NewLine +
                         " GROUP BY ServiceType, ServiceType_C " + Environment.NewLine +
                         " ORDER BY ServiceType, ServiceType_C";
            sdsCountServiceTypeB.SelectCommand = "";
            sdsCountServiceTypeB.SelectCommand = vSelectStr;
        }

        /// <summary>
        /// 客訴類別圖表用
        /// </summary>
        /// <param name="fWStr"></param>
        private void CountServiceTypeC(string fWStr)
        {
            string vSelectStr = "";
            vSelectStr = "SELECT ServiceType, ServiceType_C, COUNT(ServiceNo) AS RCount " + Environment.NewLine +
                         "  FROM ( " + Environment.NewLine +
                         "        SELECT ServiceNo, ServiceType + ISNULL(ServiceTypeB, '000') + ISNULL(ServiceTypeC, '000') AS ServiceType, " + Environment.NewLine +
                         "               (SELECT TypeText FROM CustomServiceType WHERE(TypeNo = a.ServiceType + ISNULL(a.ServiceTypeB, N'000') + ISNULL(a.ServiceTypeC, N'000'))) AS ServiceType_C " + Environment.NewLine +
                         "          FROM CustomService AS a " + Environment.NewLine +
                         "         WHERE 1 = 1 " + Environment.NewLine + fWStr +
                         "       ) AS z " + Environment.NewLine +
                         " GROUP BY ServiceType, ServiceType_C " + Environment.NewLine +
                         " ORDER BY ServiceType, ServiceType_C";
            sdsCountServiceTypeB.SelectCommand = "";
            sdsCountServiceTypeB.SelectCommand = vSelectStr;
        }

        private string GetWhereStr()
        {
            string vWStr_BuildDate = ((eCalDate_Start.Text.Trim() != "") && (eCalDate_End.Text.Trim() != "")) ?
                                     "   and a.BuildDate between '" + DateTime.Parse(eCalDate_Start.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eCalDate_Start.Text.Trim()).ToString("MM/dd") + "' and '" + DateTime.Parse(eCalDate_End.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eCalDate_End.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                     ((eCalDate_Start.Text.Trim() != "") && (eCalDate_End.Text.Trim() == "")) ?
                                     "   and a.BuildDate = '" + DateTime.Parse(eCalDate_Start.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eCalDate_Start.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                     ((eCalDate_Start.Text.Trim() == "") && (eCalDate_End.Text.Trim() != "")) ?
                                     "   and a.BuildDate = '" + DateTime.Parse(eCalDate_End.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eCalDate_End.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine : "";
            string vWStr_ServiceType = (ddlServiceType.SelectedValue.Trim() != "") ? "   and a.ServiceType = '" + ddlServiceType.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_IsClosed = (rbSelectType.SelectedValue.Trim() != "2") ? "   and isnull(IsClosed, 0) = " + rbSelectType.SelectedValue.Trim() + " " : "";
            string vResultStr = vWStr_BuildDate + vWStr_IsClosed + vWStr_ServiceType;
            return vResultStr;
        }

        /// <summary>
        /// 報表預覽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPreview_Click(object sender, EventArgs e)
        {
            //先組成 Where 語法
            string vWStr = GetWhereStr();
            switch (rbGroupBy.SelectedValue)
            {
                case "byDepNo":
                    PreviewByDepNo(vWStr);
                    break;
                case "byDriver":
                    PreviewByDriver(vWStr);
                    break;
                case "byServiceType":
                    PreviewByType(vWStr);
                    break;
                case "byLinesNo":
                    PreviewByLinesNo(vWStr);
                    break;
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
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
            string vFileName = "0800客服專線" + rbGroupBy.SelectedItem.Text.Trim() + "統計清冊";

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                //先組成 Where 語法
                string vWStr = GetWhereStr();
                //依據不同統計方式取回查詢字串
                string vSelStr = (rbGroupBy.SelectedValue == "byDepNo") ? GetSelectStr_DepNo(vWStr) :
                                 (rbGroupBy.SelectedValue == "byDriver") ? GetSelectStr_Driver(vWStr) :
                                 (rbGroupBy.SelectedValue == "byServiceType") ? GetSelectStr_Type(vWStr) :
                                 (rbGroupBy.SelectedValue == "byLinesNo") ? GetSelectStr_LinesNo(vWStr) : "";
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = drExcel.GetName(i);
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
                            if ((drExcel.GetName(i) == "筆數") && (drExcel[i].ToString() != ""))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
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
                    catch (Exception ex)
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('" + ex.Message + ex.ToString() + "')");
                        Response.Write("</" + "Script>");
                    }
                }
            }
        }
    }
}