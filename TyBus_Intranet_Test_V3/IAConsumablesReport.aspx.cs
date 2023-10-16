using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace TyBus_Intranet_Test_V3
{
    public partial class IAConsumablesReport : System.Web.UI.Page
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
                //vComputerName = Environment.MachineName; //2021.09.27 新增取得電腦名稱
                vComputerName = Page.Request.UserHostName; //2021.10.08 改變取得電腦名稱的方法

                //定義畫面上事件日期欄位各自的 onclick 事件
                string vCaseDateURL = "InputDate.aspx?TextboxID=" + eStartDate_Search.ClientID;
                string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                eStartDate_Search.Attributes["onClick"] = vCaseDateScript;
                vCaseDateURL = "InputDate.aspx?TextboxID=" + eEndDate_Search.ClientID;
                vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                eEndDate_Search.Attributes["onClick"] = vCaseDateScript;

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        plSearch.Visible = true;
                        plReport.Visible = false;
                        ddlReportName.SelectedIndex = 0;
                        SetInputState(ddlReportName.SelectedValue);
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

        private void SetInputState(string fSelectedMode)
        {
            eStartDate_Search.Text = "";
            eEndDate_Search.Text = "";
            eConsNo_S_Search.Text = "";
            eConsNo_E_Search.Text = "";
            eBrand_Search.SelectedIndex = 0;
            eCorrespondModel_Search.Text = "";
            eSystematics_Search.SelectedIndex = 0;
            eConsName_Search.Text = "";
            eConsStatus_Search.SelectedIndex = 0;
            switch (fSelectedMode)
            {
                case "P01":
                    eStartDate_Search.Enabled = false;
                    eEndDate_Search.Enabled = false;
                    eConsNo_S_Search.Enabled = true;
                    eConsNo_E_Search.Enabled = true;
                    eBrand_Search.Enabled = true;
                    eCorrespondModel_Search.Enabled = true;
                    eSystematics_Search.Enabled = true;
                    eConsName_Search.Enabled = true;
                    eConsStatus_Search.Enabled = true;
                    break;

                case "P02":
                case "P03":
                    eStartDate_Search.Enabled = true;
                    eEndDate_Search.Enabled = true;
                    eConsNo_S_Search.Enabled = true;
                    eConsNo_E_Search.Enabled = true;
                    eBrand_Search.Enabled = true;
                    eCorrespondModel_Search.Enabled = true;
                    eSystematics_Search.Enabled = true;
                    eConsName_Search.Enabled = true;
                    eConsStatus_Search.Enabled = true;
                    break;

                case "P04":
                    eStartDate_Search.Enabled = true;
                    eEndDate_Search.Enabled = true;
                    eConsNo_S_Search.Enabled = false;
                    eConsNo_E_Search.Enabled = false;
                    eBrand_Search.Enabled = false;
                    eCorrespondModel_Search.Enabled = false;
                    eSystematics_Search.Enabled = false;
                    eConsName_Search.Enabled = false;
                    eConsStatus_Search.Enabled = false;
                    break;
            }
        }

        private string GetSelectStr_P01()
        {
            string vResultStr = "";
            string vWStr_ConsNo = ((eConsNo_S_Search.Text.Trim() != "") && (eConsNo_E_Search.Text.Trim() != "")) ?
                                  "   and IA.ConsNo between '" + eConsNo_S_Search.Text.Trim() + "' and '" + eConsNo_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                  ((eConsNo_S_Search.Text.Trim() != "") && (eConsNo_E_Search.Text.Trim() == "")) ?
                                  "   and IA.ConsNo like '" + eConsNo_S_Search.Text.Trim() + "%' " + Environment.NewLine :
                                  ((eConsNo_S_Search.Text.Trim() == "") && (eConsNo_E_Search.Text.Trim() != "")) ?
                                  "   and IA.ConsNo like '" + eConsNo_E_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_Brand = (eBrand_Search.SelectedValue.Trim() != "") ? "   and IA.Brand = '" + eBrand_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CorrespondModel = (eCorrespondModel_Search.Text.Trim() != "") ? "   and IA.CorrespondModel like '%" + eCorrespondModel_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_Systematics = (eSystematics_Search.SelectedValue.Trim() != "") ? "   and IA.Systematics = '" + eSystematics_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ConsName = (eConsName_Search.Text.Trim() != "") ? "   and IA.ConsName like '%" + eConsName_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_ConsStatus = (eConsStatus_Search.SelectedValue == "S01") ? "   and IA.Quantity > IA.SaveQty " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S02") ? "   and IA.Quantity < IA.SaveQty " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S03") ? "   and IA.Quantity = 0 " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S04") ? "   and IA.IsInorder = 1 " + Environment.NewLine : "";
            vResultStr = "SELECT IA.ConsNo, (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Systematics' and ClassNo = IA.Systematics) as Systematics_C, " + Environment.NewLine +
                         "       (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Brand' and ClassNo = IA.Brand) as Brand_C, " + Environment.NewLine +
                         "       IA.ConsName, IA.OriModelNo, IA.Quantity, " + Environment.NewLine +
                         "       (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Unit' and ClassNo = IA.Unit) as Unit_C, " + Environment.NewLine +
                         "       IA.SaveQty, IA.Position, IA.Spec_Color, IA.Spec_Other, IA.CorrespondModel, IA.LastIndate, IA.LastInQty, " + Environment.NewLine +
                         "       case when isnull(IA.IsInorder, 0) = 1 then 'V' else 'X' end IsInorder, " + Environment.NewLine +
                         "       case when isnull(IA.IsStopUse, 0) = 1 then 'V' else 'X' end IsStopUse, " + Environment.NewLine +
                         "       IA.Remark, IA.BuMan, (select[Name] from Employee where EmpNo = IA.BuMan) as BuManName, IA.BuDate, " + Environment.NewLine +
                         "       IA.ModifyMan, (select[Name] from Employee where EmpNo = IA.ModifyMan) as ModifyManName, IA.ModifyDate, IA.Price " + Environment.NewLine +
                         "  FROM IAConsumables as IA " + Environment.NewLine +
                         " WHERE isnull(IA.ConsNo, '') <> '' " + Environment.NewLine +
                         vWStr_ConsNo +
                         vWStr_Brand +
                         vWStr_CorrespondModel +
                         vWStr_Systematics +
                         vWStr_ConsName +
                         vWStr_ConsStatus +
                         " Order by ConsNo DESC";
            return vResultStr;
        }

        private string GetSelectStr_P02()
        {
            string vResultStr = "";
            string vWStr_BuDate = ((eStartDate_Search.Text.Trim() != "") && (eEndDate_Search.Text.Trim() != "")) ?
                                  "   and tA.BuDate between '" + eStartDate_Search.Text.Trim() + "' and '" + eEndDate_Search.Text.Trim() + "'" + Environment.NewLine :
                                  ((eStartDate_Search.Text.Trim() != "") && (eEndDate_Search.Text.Trim() == "")) ?
                                  "   and tA.BuDate >= '" + eStartDate_Search.Text.Trim() + "' " + Environment.NewLine :
                                  ((eStartDate_Search.Text.Trim() == "") && (eEndDate_Search.Text.Trim() != "")) ?
                                  "   and tA.BuDate <= '" + eEndDate_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ConsNo = ((eConsNo_S_Search.Text.Trim() != "") && (eConsNo_E_Search.Text.Trim() != "")) ?
                                  "   and tB.ConsNo between '" + eConsNo_S_Search.Text.Trim() + "' and '" + eConsNo_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                  ((eConsNo_S_Search.Text.Trim() != "") && (eConsNo_E_Search.Text.Trim() == "")) ?
                                  "   and tB.ConsNo like '" + eConsNo_S_Search.Text.Trim() + "%' " + Environment.NewLine :
                                  ((eConsNo_S_Search.Text.Trim() == "") && (eConsNo_E_Search.Text.Trim() != "")) ?
                                  "   and tB.ConsNo like '" + eConsNo_E_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_Brand = (eBrand_Search.SelectedValue.Trim() != "") ? "   and tC.Brand = '" + eBrand_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CorrespondModel = (eCorrespondModel_Search.Text.Trim() != "") ? "   and tC.CorrespondModel like '%" + eCorrespondModel_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_Systematics = (eSystematics_Search.SelectedValue.Trim() != "") ? "   and tC.Systematics = '" + eSystematics_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ConsName = (eConsName_Search.Text.Trim() != "") ? "   and tC.ConsName like '%" + eConsName_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_ConsStatus = (eConsStatus_Search.SelectedValue == "S01") ? "   and tC.Quantity > tC.SaveQty " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S02") ? "   and tC.Quantity < tC.SaveQty " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S03") ? "   and tC.Quantity = 0 " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S04") ? "   and tC.IsInorder = 1 " + Environment.NewLine : "";
            vResultStr = "select tB.ConsNo, tC.ConsName, tC.Brand, tC.OriModelNo, tC.Spec_Color, tC.CorrespondModel, tC.Quantity, max(tA.BuDate) LastInDate, sum(tB.Quantity) as OutQty " + Environment.NewLine +
                         "  from IASheetB as tB left join IASheetA as tA on tA.SheetNo = tB.SheetNo " + Environment.NewLine +
                         "                      left join IAConsumables as tC on tC.ConsNo = tB.ConsNo " + Environment.NewLine +
                         " where tB.QtyMode = 1 " + Environment.NewLine +
                         vWStr_BuDate +
                         vWStr_ConsNo +
                         vWStr_Brand +
                         vWStr_CorrespondModel +
                         vWStr_Systematics +
                         vWStr_ConsName +
                         vWStr_ConsStatus +
                         " group by tB.ConsNo, tC.ConsName, tC.Brand, tC.OriModelNo, tC.Spec_Color, tC.CorrespondModel, tC.Quantity " + Environment.NewLine +
                         " order by tB.ConsNo";
            return vResultStr;
        }

        private string GetSelectStr_P03()
        {
            string vResultStr = "";
            string vWStr_BuDate = ((eStartDate_Search.Text.Trim() != "") && (eEndDate_Search.Text.Trim() != "")) ?
                                  "   and tA.BuDate between '" + eStartDate_Search.Text.Trim() + "' and '" + eEndDate_Search.Text.Trim() + "'" + Environment.NewLine :
                                  ((eStartDate_Search.Text.Trim() != "") && (eEndDate_Search.Text.Trim() == "")) ?
                                  "   and tA.BuDate >= '" + eStartDate_Search.Text.Trim() + "' " + Environment.NewLine :
                                  ((eStartDate_Search.Text.Trim() == "") && (eEndDate_Search.Text.Trim() != "")) ?
                                  "   and tA.BuDate <= '" + eEndDate_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ConsNo = ((eConsNo_S_Search.Text.Trim() != "") && (eConsNo_E_Search.Text.Trim() != "")) ?
                                  "   and tB.ConsNo between '" + eConsNo_S_Search.Text.Trim() + "' and '" + eConsNo_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                  ((eConsNo_S_Search.Text.Trim() != "") && (eConsNo_E_Search.Text.Trim() == "")) ?
                                  "   and tB.ConsNo like '" + eConsNo_S_Search.Text.Trim() + "%' " + Environment.NewLine :
                                  ((eConsNo_S_Search.Text.Trim() == "") && (eConsNo_E_Search.Text.Trim() != "")) ?
                                  "   and tB.ConsNo like '" + eConsNo_E_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_Brand = (eBrand_Search.SelectedValue.Trim() != "") ? "   and tC.Brand = '" + eBrand_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CorrespondModel = (eCorrespondModel_Search.Text.Trim() != "") ? "   and tC.CorrespondModel like '%" + eCorrespondModel_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_Systematics = (eSystematics_Search.SelectedValue.Trim() != "") ? "   and tC.Systematics = '" + eSystematics_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ConsName = (eConsName_Search.Text.Trim() != "") ? "   and tC.ConsName like '%" + eConsName_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_ConsStatus = (eConsStatus_Search.SelectedValue == "S01") ? "   and tC.Quantity > tC.SaveQty " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S02") ? "   and tC.Quantity < tC.SaveQty " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S03") ? "   and tC.Quantity = 0 " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S04") ? "   and tC.IsInorder = 1 " + Environment.NewLine : "";
            vResultStr = "select tB.ConsNo, tC.ConsName, tC.Brand, tC.OriModelNo, tC.Spec_Color, tC.CorrespondModel, tC.Quantity, max(tA.BuDate) LastOutDate, sum(tB.Quantity) as OutQty " + Environment.NewLine +
                         "  from IASheetB as tB left join IASheetA as tA on tA.SheetNo = tB.SheetNo " + Environment.NewLine +
                         "                      left join IAConsumables as tC on tC.ConsNo = tB.ConsNo " + Environment.NewLine +
                         " where tB.QtyMode = -1 " + Environment.NewLine +
                         vWStr_BuDate +
                         vWStr_ConsNo +
                         vWStr_Brand +
                         vWStr_CorrespondModel +
                         vWStr_Systematics +
                         vWStr_ConsName +
                         vWStr_ConsStatus +
                         " group by tB.ConsNo, tC.ConsName, tC.Brand, tC.OriModelNo, tC.Spec_Color, tC.CorrespondModel, tC.Quantity " + Environment.NewLine +
                         " order by tB.ConsNo";
            return vResultStr;
        }

        private string GetSelectStr_P04()
        {
            string vResultStr = "";
            string vWStr_BuDate = ((eStartDate_Search.Text.Trim() != "") && (eEndDate_Search.Text.Trim() != "")) ?
                                  "   and tA.BuDate between '" + eStartDate_Search.Text.Trim() + "' and '" + eEndDate_Search.Text.Trim() + "'" + Environment.NewLine :
                                  ((eStartDate_Search.Text.Trim() != "") && (eEndDate_Search.Text.Trim() == "")) ?
                                  "   and tA.BuDate >= '" + eStartDate_Search.Text.Trim() + "' " + Environment.NewLine :
                                  ((eStartDate_Search.Text.Trim() == "") && (eEndDate_Search.Text.Trim() != "")) ?
                                  "   and tA.BuDate <= '" + eEndDate_Search.Text.Trim() + "' " + Environment.NewLine : "";

            vResultStr = "select tB.ConsNo, tC.ConsName, tC.Brand, tC.OriModelNo, tC.Spec_Color, tC.CorrespondModel, tC.Quantity, tC.LastIndate, sum(tB.Quantity) as InQty " + Environment.NewLine +
                         "  from IASheetB as tB left join IASheetA as tA on tA.SheetNo = tB.SheetNo " + Environment.NewLine +
                         "                      left join IAConsumables as tC on tC.ConsNo = tB.ConsNo " + Environment.NewLine +
                         " where tB.QtyMode = -1 " + Environment.NewLine +
                         vWStr_BuDate +
                         " group by tB.ConsNo, tC.ConsName, tC.Brand, tC.OriModelNo, tC.Spec_Color, tC.CorrespondModel, tC.Quantity, tC.LastIndate " + Environment.NewLine +
                         " order by tB.ConsNo";
            return vResultStr;
        }

        private DataTable getTable(string fSelectStr)
        {
            DataTable dtResult = new DataTable();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daTemp = new SqlDataAdapter(fSelectStr, connTemp);
                connTemp.Open();
                daTemp.Fill(dtResult);
            }
            return dtResult;
        }

        protected void bbPreview_Click(object sender, EventArgs e)
        {
            string vSelectStr = "";
            DataTable dtPrint = new DataTable();
            string vCompanyName = "桃園汽車客運股份有限公司";
            string vReportName = ddlReportName.SelectedItem.Text.Trim();
            string vReportPath = "";
            string vDataSourceName = "";
            switch (ddlReportName.SelectedValue)
            {
                case "P01":
                    vSelectStr = GetSelectStr_P01();
                    dtPrint = getTable(vSelectStr);
                    vReportPath = @"Report\IAConsumablesP.rdlc";
                    vDataSourceName = "IAConsumablesP";
                    break;

                case "P02":
                    vSelectStr = GetSelectStr_P02();
                    dtPrint = getTable(vSelectStr);
                    vReportPath = @"Report\IAConsumablesP02.rdlc";
                    vDataSourceName = "IAConsumablesP02";
                    break;

                case "P03":
                    vSelectStr = GetSelectStr_P03();
                    dtPrint = getTable(vSelectStr);
                    vReportPath = @"Report\IAConsumablesP03.rdlc";
                    vDataSourceName = "IAConsumablesP03";
                    break;

                case "P04":
                    vSelectStr = GetSelectStr_P04();
                    dtPrint = getTable(vSelectStr);
                    vReportPath = @"Report\IAConsumablesP04.rdlc";
                    vDataSourceName = "IAConsumablesP04";
                    break;
            }
            if (dtPrint.Rows.Count > 0)
            {
                ReportDataSource rdsPrint = new ReportDataSource(vDataSourceName, dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = vReportPath;
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                if (eStartDate_Search.Enabled)
                {
                    rvPrint.LocalReport.SetParameters(new ReportParameter("StartDate", eStartDate_Search.Text));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("EndDate", eEndDate_Search.Text));
                }
                rvPrint.LocalReport.Refresh();
                plReport.Visible = true;
                plSearch.Visible = false;
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vSelectStr = "";
            string vFileName = ddlReportName.SelectedItem.Text.Trim();
            DateTime vBuDate;
            switch (ddlReportName.SelectedValue)
            {
                case "P01":
                    vSelectStr = GetSelectStr_P01();
                    break;

                case "P02":
                    vSelectStr = GetSelectStr_P02();
                    break;

                case "P03":
                    vSelectStr = GetSelectStr_P03();
                    break;

                case "P04":
                    vSelectStr = GetSelectStr_P04();
                    break;
            }
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
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
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
                        vHeaderText = (drExcel.GetName(i) == "ConsNo") ? "料號" :
                                      (drExcel.GetName(i) == "ConsName") ? "品名" :
                                      (drExcel.GetName(i) == "Brand") ? "廠牌" :
                                      (drExcel.GetName(i) == "Brand_C") ? "廠牌" :
                                      (drExcel.GetName(i) == "OriModelNo") ? "原廠料號" :
                                      (drExcel.GetName(i) == "Spec_Color") ? "顏色" :
                                      (drExcel.GetName(i) == "CorrespondModel") ? "適用機型" :
                                      (drExcel.GetName(i) == "BuMan") ? "建檔人工號" :
                                      (drExcel.GetName(i) == "BuManName") ? "建檔人" :
                                      (drExcel.GetName(i) == "Quantity") ? "數量" :
                                      (drExcel.GetName(i) == "LastOutDate") ? "最後出貨日" :
                                      (drExcel.GetName(i) == "OutQty") ? "出貨量" :
                                      (drExcel.GetName(i) == "LastInDate") ? "最後進貨日" :
                                      (drExcel.GetName(i) == "InQty") ? "進貨量" :
                                      (drExcel.GetName(i) == "Systematics_C") ? "類別" :
                                      (drExcel.GetName(i) == "Unit_C") ? "單位" :
                                      (drExcel.GetName(i) == "SaveQty") ? "安全量" :
                                      (drExcel.GetName(i) == "Position") ? "庫位" :
                                      (drExcel.GetName(i) == "Spec_Other") ? "其他規格" :
                                      (drExcel.GetName(i) == "IsInorder") ? "採購中" :
                                      (drExcel.GetName(i) == "IsStopUse") ? "停用" :
                                      (drExcel.GetName(i) == "Price") ? "單價" :
                                      (drExcel.GetName(i) == "Remark") ? "備註" :
                                      (drExcel.GetName(i) == "BuDate") ? "建檔日期" :
                                      (drExcel.GetName(i) == "ModifyMan") ? "異動人工號" :
                                      (drExcel.GetName(i) == "ModifyManName") ? "異動人" :
                                      (drExcel.GetName(i) == "ModifyDate") ? "異動日期" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i)=="LastOutDate")&&(drExcel[i].ToString().Trim()!=""))||
                                ((drExcel.GetName(i) == "LastInDate") && (drExcel[i].ToString().Trim() != ""))||
                                ((drExcel.GetName(i) == "BuDate") && (drExcel[i].ToString().Trim() != ""))||
                                ((drExcel.GetName(i) == "ModifyDate") && (drExcel[i].ToString().Trim() != "")))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if (((drExcel.GetName(i) == "Quantity") && (drExcel[i].ToString().Trim() != "")) ||
                                     ((drExcel.GetName(i) == "SaveQty") && (drExcel[i].ToString().Trim() != "")) ||
                                     ((drExcel.GetName(i) == "Price") && (drExcel[i].ToString().Trim() != "")))
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

                        if (msTarget.Length>0)
                        {
                            //寫入操作記錄
                            string vRecord_BuDate = ((eStartDate_Search.Text.Trim() != "") && (eEndDate_Search.Text.Trim() != "")) ?
                                                    " ["+eStartDate_Search.Text.Trim() + "] ～ [" + eEndDate_Search.Text.Trim() + "]" + Environment.NewLine :
                                                    ((eStartDate_Search.Text.Trim() != "") && (eEndDate_Search.Text.Trim() == "")) ?
                                                    " [" + eStartDate_Search.Text.Trim() + "] " + Environment.NewLine :
                                                    ((eStartDate_Search.Text.Trim() == "") && (eEndDate_Search.Text.Trim() != "")) ?
                                                    " [" + eEndDate_Search.Text.Trim() + "] " + Environment.NewLine : "";
                            string vRecord_ConsNo = ((eConsNo_S_Search.Text.Trim() != "") && (eConsNo_E_Search.Text.Trim() != "")) ?
                                                  " [" + eConsNo_S_Search.Text.Trim() + "] ～ [" + eConsNo_E_Search.Text.Trim() + "] " + Environment.NewLine :
                                                  ((eConsNo_S_Search.Text.Trim() != "") && (eConsNo_E_Search.Text.Trim() == "")) ?
                                                  " [" + eConsNo_S_Search.Text.Trim() + "] " + Environment.NewLine :
                                                  ((eConsNo_S_Search.Text.Trim() == "") && (eConsNo_E_Search.Text.Trim() != "")) ?
                                                  " [" + eConsNo_E_Search.Text.Trim() + "] " + Environment.NewLine : "";
                            string vRecord_Brand = (eBrand_Search.SelectedValue.Trim() != "") ? " [" + eBrand_Search.SelectedValue.Trim() + "] " + Environment.NewLine : "";
                            string vRecord_CorrespondModel = (eCorrespondModel_Search.Text.Trim() != "") ? " [" + eCorrespondModel_Search.Text.Trim() + "] " + Environment.NewLine : "";
                            string vRecord_Systematics = (eSystematics_Search.SelectedValue.Trim() != "") ? " [" + eSystematics_Search.SelectedValue.Trim() + "] " + Environment.NewLine : "";
                            string vRecord_ConsName = (eConsName_Search.Text.Trim() != "") ? " [" + eConsName_Search.Text.Trim() + "] " + Environment.NewLine : "";
                            string vRecord_ConsStatus = eConsStatus_Search.SelectedItem.Text.Trim();
                            string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                 "IAConsumblesReport.aspx" + Environment.NewLine +
                                                 "建檔日期：" + vRecord_BuDate + Environment.NewLine +
                                                 "料號：" + vRecord_ConsNo + Environment.NewLine +
                                                 "廠牌：" + vRecord_Brand + Environment.NewLine +
                                                 "適用型號：" + vRecord_CorrespondModel + Environment.NewLine +
                                                 "耗材類別：" + vRecord_Systematics + Environment.NewLine +
                                                 "品名字樣：" + vRecord_ConsName + Environment.NewLine +
                                                 "庫存狀況：" + vRecord_ConsStatus;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);

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
                    catch (Exception eMessage)
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('" + eMessage.Message + "')");
                        Response.Write("</" + "Script>");
                    }
                }
            }
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plSearch.Visible = true;
            plReport.Visible = false;
        }

        protected void ddlReportName_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetInputState(ddlReportName.SelectedValue);
        }
    }
}