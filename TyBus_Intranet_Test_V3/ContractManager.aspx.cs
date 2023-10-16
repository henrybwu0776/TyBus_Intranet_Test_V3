using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class ContractManager : System.Web.UI.Page
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
        private string vSQLStr_Main = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"] != null)
            {
                //顯示民國年
                CultureInfo Cal = new CultureInfo("zh-TW");
                Cal.DateTimeFormat.Calendar = new TaiwanCalendar();
                Cal.DateTimeFormat.YearMonthPattern = "民國 yyy 年 MM 月";
                System.Threading.Thread.CurrentThread.CurrentCulture = Cal; ;
                //
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

                if (vLoginID != "")
                {
                    if (!IsPostBack)
                    {
                        if ((vLoginDepNo == "02") || (vLoginDepNo == "03"))
                        {
                            vSQLStr_Main = "select cast('' as varchar) DepNo, cast('' as varchar) DepName " + Environment.NewLine +
                                           " union all " + Environment.NewLine +
                                           "select DepNo, [Name] as DepName " + Environment.NewLine +
                                           "  from Department " + Environment.NewLine +
                                           " where DepNo in (select distinct DepNo from Employee where isnull(LeaveDay, '9999/12/31') = '9999/12/31')" + Environment.NewLine +
                                           "   and DepNo > '00' and DepNo < '80' ";
                        }
                        else
                        {
                            vSQLStr_Main = "select cast('' as varchar) DepNo, cast('' as varchar) DepName " + Environment.NewLine +
                                           " union all " + Environment.NewLine +
                                           "select DepNo, [Name] as DepName " + Environment.NewLine +
                                           "  from Department " + Environment.NewLine +
                                           " where DepNo ='" + vLoginDepNo + "' ";
                        }
                        sdsResponsibleDepList.SelectCommand = vSQLStr_Main;
                        sdsWorkDepList.SelectCommand = vSQLStr_Main;
                        ddlWorkDep_Search.DataBind();
                        ddlResponsibleDep_Search.DataBind();
                    }
                    else
                    {
                        OpenData();
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

        private string GetSelectStr()
        {
            string vWStr_Dep = "";
            if ((vLoginDepNo == "02") || (vLoginDepNo == "03"))
            {
                vWStr_Dep = ((ddlWorkDep_Search.SelectedValue.Trim() != "") && (ddlResponsibleDep_Search.SelectedValue.Trim() == "")) ?
                            "   and cm.WorkDep = '" + ddlWorkDep_Search.SelectedValue.Trim() + "' " + Environment.NewLine :
                            ((ddlWorkDep_Search.SelectedValue.Trim() == "") && (ddlResponsibleDep_Search.SelectedValue.Trim() != "")) ?
                            "   and cm.ResponsibleDep = '" + ddlResponsibleDep_Search.SelectedValue.Trim() + "' " + Environment.NewLine :
                            ((ddlWorkDep_Search.SelectedValue.Trim() != "") && (ddlResponsibleDep_Search.SelectedValue.Trim() != "")) ?
                            "   and cm.WorkDep = '" + ddlWorkDep_Search.SelectedValue.Trim() + "' and cm.ResponsibleDep = '" + ddlResponsibleDep_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            }
            else
            {
                vWStr_Dep = ((ddlWorkDep_Search.SelectedValue.Trim() == "") && (ddlResponsibleDep_Search.SelectedValue.Trim() == "")) ?
                            "   and cm.ResponsibleDep = '" + vLoginDepNo + "' " + Environment.NewLine :
                            ((ddlWorkDep_Search.SelectedValue.Trim() != "") && (ddlResponsibleDep_Search.SelectedValue.Trim() == "")) ?
                            "   and cm.WorkDep = '" + ddlWorkDep_Search.SelectedValue.Trim() + "' " + Environment.NewLine :
                            ((ddlWorkDep_Search.SelectedValue.Trim() == "") && (ddlResponsibleDep_Search.SelectedValue.Trim() != "")) ?
                            "   and cm.ResponsibleDep = '" + ddlResponsibleDep_Search.SelectedValue.Trim() + "' " + Environment.NewLine :
                            ((ddlWorkDep_Search.SelectedValue.Trim() != "") && (ddlResponsibleDep_Search.SelectedValue.Trim() != "")) ?
                            "   and cm.WorkDep = '" + ddlWorkDep_Search.SelectedValue.Trim() + "' and cm.ResponsibleDep = '" + ddlResponsibleDep_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            }
            string vWStr_CustomName = (eCustomName_Search.Text.Trim() != "") ? "   and cm.CustomName like '%" + eCustomName_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vResultStr = "select IndexNo, ContractNo, " + Environment.NewLine +
                                "       WorkDep, (SELECT NAME FROM DEPARTMENT WHERE DEPNO = cm.WorkDep) AS WorkDepName, " + Environment.NewLine +
                                "       ResponsibleDep, (SELECT NAME FROM DEPARTMENT WHERE DEPNO = cm.ResponsibleDep) AS ResponsibleDepName, " + Environment.NewLine +
                                "       InvoiceNo, CustomName, ContractText " + Environment.NewLine +
                                "  from ContractManager cm " + Environment.NewLine +
                                " where isnull(cm.IndexNo, '') <> '' " + Environment.NewLine +
                                vWStr_Dep + vWStr_CustomName +
                                " order by IndexNo ";
            return vResultStr;
        }

        private void OpenData()
        {
            sdsContractManagerList.SelectCommand = GetSelectStr();
            gridContractManager.DataBind();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        private string GetSelectStr_Excel()
        {
            string vWStr_Dep = "";
            if ((vLoginDepNo == "02") || (vLoginDepNo == "03"))
            {
                vWStr_Dep = ((ddlWorkDep_Search.SelectedValue.Trim() != "") && (ddlResponsibleDep_Search.SelectedValue.Trim() == "")) ?
                            "   and cm.WorkDep = '" + ddlWorkDep_Search.SelectedValue.Trim() + "' " + Environment.NewLine :
                            ((ddlWorkDep_Search.SelectedValue.Trim() == "") && (ddlResponsibleDep_Search.SelectedValue.Trim() != "")) ?
                            "   and cm.ResponsibleDep = '" + ddlResponsibleDep_Search.SelectedValue.Trim() + "' " + Environment.NewLine :
                            ((ddlWorkDep_Search.SelectedValue.Trim() != "") && (ddlResponsibleDep_Search.SelectedValue.Trim() != "")) ?
                            "   and cm.WorkDep = '" + ddlWorkDep_Search.SelectedValue.Trim() + "' and cm.ResponsibleDep = '" + ddlResponsibleDep_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            }
            else
            {
                vWStr_Dep = ((ddlWorkDep_Search.SelectedValue.Trim() == "") && (ddlResponsibleDep_Search.SelectedValue.Trim() == "")) ?
                            "   and cm.ResponsibleDep = '" + vLoginDepNo + "' " + Environment.NewLine :
                            ((ddlWorkDep_Search.SelectedValue.Trim() != "") && (ddlResponsibleDep_Search.SelectedValue.Trim() == "")) ?
                            "   and cm.WorkDep = '" + ddlWorkDep_Search.SelectedValue.Trim() + "' " + Environment.NewLine :
                            ((ddlWorkDep_Search.SelectedValue.Trim() == "") && (ddlResponsibleDep_Search.SelectedValue.Trim() != "")) ?
                            "   and cm.ResponsibleDep = '" + ddlResponsibleDep_Search.SelectedValue.Trim() + "' " + Environment.NewLine :
                            ((ddlWorkDep_Search.SelectedValue.Trim() != "") && (ddlResponsibleDep_Search.SelectedValue.Trim() != "")) ?
                            "   and cm.WorkDep = '" + ddlWorkDep_Search.SelectedValue.Trim() + "' and cm.ResponsibleDep = '" + ddlResponsibleDep_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            }
            string vWStr_CustomName = (eCustomName_Search.Text.Trim() != "") ? "   and cm.CustomName like '%" + eCustomName_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vResultStr = "select (SELECT NAME FROM DEPARTMENT WHERE DEPNO = cm.WorkDep) AS WorkDepName, " + Environment.NewLine +
                                "       (SELECT NAME FROM DEPARTMENT WHERE DEPNO = cm.ResponsibleDep) AS ResponsibleDepName, " + Environment.NewLine +
                                "       ContractNo, InvoiceNo, CustomName, ContractText, AssignDate, EffectiveDateS, EffectiveDateE, " + Environment.NewLine +
                                "       IsClose, Amount, Tax, TotalAmount, StampDuty, Remark " + Environment.NewLine +
                                "  from ContractManager cm " + Environment.NewLine +
                                " where isnull(cm.IndexNo, '') <> '' " + Environment.NewLine +
                                vWStr_Dep + vWStr_CustomName +
                                " order by ContractNo ";
            return vResultStr;
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

            HSSFCellStyle csData_Float = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csData_Int.Alignment = HorizontalAlignment.Right;

            HSSFDataFormat format_Float = (HSSFDataFormat)wbExcel.CreateDataFormat();
            csData_Float.DataFormat = format.GetFormat("##0.00");

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                string vSelStr_Excel = GetSelectStr_Excel();
                SqlCommand cmdExcel = new SqlCommand(vSelStr_Excel, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //有資料才做下面的動作
                    string vFileName = "合約書清冊";
                    string vHeaderText = "";
                    int vLinesNo = 0;
                    DateTime vBuDate;
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "WORKDEPNAME") ? "承辦單位" :
                                      (drExcel.GetName(i).ToUpper() == "RESPONSIBLEDEPNAME") ? "主責單位" :
                                      (drExcel.GetName(i).ToUpper() == "CONTRACTNO") ? "編號" :
                                      (drExcel.GetName(i).ToUpper() == "INVOICENO") ? "對方統一編號" :
                                      (drExcel.GetName(i).ToUpper() == "CUSTOMNAME") ? "對方名稱" :
                                      (drExcel.GetName(i).ToUpper() == "CONTRACTTEXT") ? "工程地點、內容或名稱" :
                                      (drExcel.GetName(i).ToUpper() == "ASSIGNDATE") ? "立契(據)日期" :
                                      (drExcel.GetName(i).ToUpper() == "EFFECTIVEDATES") ? "合約書期間(起)" :
                                      (drExcel.GetName(i).ToUpper() == "EFFECTIVEDATEE") ? "合約書期間(迄)" :
                                      (drExcel.GetName(i).ToUpper() == "ISCLOSE") ? "是否結案" :
                                      (drExcel.GetName(i).ToUpper() == "AMOUNT") ? "憑證金額" :
                                      (drExcel.GetName(i).ToUpper() == "TAX") ? "稅額" :
                                      (drExcel.GetName(i).ToUpper() == "TOTALAMOUNT") ? "合計" :
                                      (drExcel.GetName(i).ToUpper() == "STAMPDUTY") ? "印花稅" :
                                      (drExcel.GetName(i).ToUpper() == "REMARK") ? "備註" : drExcel.GetName(i);
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    while (drExcel.Read())
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            if (((drExcel.GetName(i).ToUpper() == "ASSIGNDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "EFFECTIVEDATES") ||
                                 (drExcel.GetName(i).ToUpper() == "EFFECTIVEDATEE")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue((vBuDate.Year - 1911).ToString() + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if (((drExcel.GetName(i).ToUpper() == "AMOUNT") ||
                                      (drExcel.GetName(i).ToUpper() == "TAX") ||
                                      (drExcel.GetName(i).ToUpper() == "TOTALAMOUNT") ||
                                      (drExcel.GetName(i).ToUpper() == "STAMPDUTY")) && (drExcel[i].ToString() != ""))
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
                            string vDepNo_Main = (ddlWorkDep_Search.SelectedValue.Trim() != "") ? ddlWorkDep_Search.SelectedValue : "全部";
                            string vDepNo_Respond = (ddlResponsibleDep_Search.SelectedValue.Trim() != "") ? ddlResponsibleDep_Search.SelectedValue : "全部";
                            string vCustName = (eCustomName_Search.Text.Trim() != "") ? eCustomName_Search.Text.Trim() : "";
                            string vRecordNote = "匯出檔案_合約書管理" + Environment.NewLine +
                                                 "ContractManager.aspx" + Environment.NewLine +
                                                 "承辦單位：" + vDepNo_Main +
                                                 "主責單位：" + vDepNo_Respond +
                                                 "對方名稱" + vCustName;
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

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            Label eIndexNo_Edit = (Label)fvContractManager.FindControl("eIndexNo_Edit");
            if (eIndexNo_Edit != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vIndexNo = eIndexNo_Edit.Text.Trim();
                try
                {
                    //先複製一份到異動記錄檔
                    string vTempStr = "select max(Items) MaxItem from ContractManager_History where IndexNo = '" + vIndexNo + "' ";
                    string vMaxItemsStr = PF.GetValue(vConnStr, vTempStr, "MaxItem");
                    int vMaxItem = (vMaxItemsStr != "") ? Int32.Parse(vMaxItemsStr) : 0;
                    string vNewItems = (vMaxItem + 1).ToString("D4");
                    string vIndexNoItem = vIndexNo.Trim() + vNewItems.Trim();
                    vTempStr = "INSERT INTO ContractManager_History " + Environment.NewLine +
                               "       (IndexNoItem, Items, ModifyMode, IndexNo, ContractNo, WorkDep, ResponsibleDep, InvoiceNo, CustomName, ContractText, " + Environment.NewLine +
                               "        AssignDate, EffectiveDateS, EffectiveDateE, IsClose, Amount, Tax, TotalAmount, StampDuty, Remark, BuMan, BuDate) " + Environment.NewLine +
                               "SELECT '" + vIndexNoItem + "', '" + vNewItems + "', 'EDIT', IndexNo, ContractNo, WorkDep, ResponsibleDep, InvoiceNo, " + Environment.NewLine +
                               "       CustomName, ContractText, AssignDate, EffectiveDateS, EffectiveDateE, IsClose, Amount, Tax, TotalAmount, " + Environment.NewLine +
                               "       StampDuty, Remark, '" + vLoginID + "', GetDate() " + Environment.NewLine +
                               "  FROM ContractManager " + Environment.NewLine +
                               " WHERE IndexNo = '" + vIndexNo + "' ";
                    PF.ExecSQL(vConnStr, vTempStr);
                    //再來進行實際的更新資料
                    TextBox eContractNo_Edit = (TextBox)fvContractManager.FindControl("eContractNo_Edit");
                    string vContractNo = eContractNo_Edit.Text.Trim();
                    TextBox eWorkDep_Edit = (TextBox)fvContractManager.FindControl("eWorkDep_Edit");
                    string vWorkDep = eWorkDep_Edit.Text.Trim();
                    TextBox eResponsibleDep_Edit = (TextBox)fvContractManager.FindControl("eResponsibleDep_Edit");
                    string vResponsibleDep = eResponsibleDep_Edit.Text.Trim();
                    TextBox eInvoiceNo_Edit = (TextBox)fvContractManager.FindControl("eInvoiceNo_Edit");
                    string vInvoiceNo = eInvoiceNo_Edit.Text.Trim();
                    TextBox eCustomName_Edit = (TextBox)fvContractManager.FindControl("eCustomName_Edit");
                    string vCustomName = eCustomName_Edit.Text.Trim();
                    TextBox eContractText_Edit = (TextBox)fvContractManager.FindControl("eContractText_Edit");
                    string vContractText = eContractText_Edit.Text.Trim();
                    TextBox eAssignDate_Edit = (TextBox)fvContractManager.FindControl("eAssignDate_Edit");
                    string vAssignDateStr = eAssignDate_Edit.Text.Trim();
                    DateTime vAssignDate;
                    if (!DateTime.TryParse(vAssignDateStr, out vAssignDate))
                    {
                        vAssignDate = DateTime.Today;
                    }
                    TextBox eEffectiveDateS_Edit = (TextBox)fvContractManager.FindControl("eEffectiveDateS_Edit");
                    string vEffectiveDateSStr = eEffectiveDateS_Edit.Text.Trim();
                    DateTime vEffectiveDateS;
                    if (!DateTime.TryParse(vEffectiveDateSStr, out vEffectiveDateS))
                    {
                        vEffectiveDateS = DateTime.Today;
                    }
                    TextBox eEffectiveDateE_Edit = (TextBox)fvContractManager.FindControl("eEffectiveDateE_Edit");
                    string vEffectiveDateEStr = eEffectiveDateE_Edit.Text.Trim();
                    DateTime vEffectiveDateE;
                    if (!DateTime.TryParse(vEffectiveDateEStr, out vEffectiveDateE))
                    {
                        vEffectiveDateE = DateTime.Today;
                    }
                    Label eIsClose_Edit = (Label)fvContractManager.FindControl("eIsClose_Edit");
                    string vIsClose = (eIsClose_Edit.Text.Trim() != "") ? eIsClose_Edit.Text.Trim() : "X";
                    TextBox eAmount_Edit = (TextBox)fvContractManager.FindControl("eAmount_Edit");
                    string vAmountStr = eAmount_Edit.Text.Trim();
                    TextBox eTax_Edit = (TextBox)fvContractManager.FindControl("eTax_Edit");
                    string vTaxStr = eTax_Edit.Text.Trim();
                    Label eTotalAmount_Edit = (Label)fvContractManager.FindControl("eTotalAmount_Edit");
                    string vTotalAmountStr = eTotalAmount_Edit.Text.Trim();
                    TextBox eStampDuty_Edit = (TextBox)fvContractManager.FindControl("eStampDuty_Edit");
                    string vStampDutyStr = eStampDuty_Edit.Text.Trim();
                    TextBox eRemark_Edit = (TextBox)fvContractManager.FindControl("eRemark_Edit");
                    string vRemark = eRemark_Edit.Text.Trim();

                    sdsContractManagerDetail.UpdateCommand = "update ContractManager " + Environment.NewLine +
                                                             "   set ContractNo = @ContractNo, WorkDep = @WorkDep, ResponsibleDep = @ResponsibleDep, " + Environment.NewLine +
                                                             "       InvoiceNo = @InvoiceNo, CustomName = @CustomName, ContractText = @ContractText, " + Environment.NewLine +
                                                             "       AssignDate = @AssignDate, EffectiveDateS = @EffectiveDateS, EffectiveDateE = @EffectiveDateE, " + Environment.NewLine +
                                                             "       IsClose = @IsClose, Amount = @Amount, Tax = @Tax, TotalAmount = @TotalAmount, StampDuty = @StampDuty, " + Environment.NewLine +
                                                             "       Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate " + Environment.NewLine +
                                                             " where IndexNo = @IndexNo ";
                    sdsContractManagerDetail.UpdateParameters.Clear();
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("ContractNo", DbType.String, (vContractNo != "") ? vContractNo : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("WorkDep", DbType.String, (vWorkDep != "") ? vWorkDep : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("ResponsibleDep", DbType.String, (vResponsibleDep != "") ? vResponsibleDep : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("InvoiceNo", DbType.String, (vInvoiceNo != "") ? vInvoiceNo : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("CustomName", DbType.String, (vCustomName != "") ? vCustomName : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("ContractText", DbType.String, (vContractText != "") ? vContractText : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("AssignDate", DbType.Date, (vAssignDateStr != "") ? vAssignDate.ToShortDateString() : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("EffectiveDateS", DbType.Date, (vEffectiveDateSStr != "") ? vEffectiveDateS.ToShortDateString() : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("EffectiveDateE", DbType.Date, (vEffectiveDateEStr != "") ? vEffectiveDateE.ToShortDateString() : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("IsClose", DbType.String, (vIsClose != "") ? vIsClose : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("Amount", DbType.Int32, (vAmountStr != "") ? vAmountStr : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("Tax", DbType.Int32, (vTaxStr != "") ? vTaxStr : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Int32, (vTotalAmountStr != "") ? vTotalAmountStr : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("StampDuty", DbType.Int32, (vStampDutyStr != "") ? vStampDutyStr : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("Remark", DbType.String, (vRemark != "") ? vRemark : String.Empty));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("ModifyDate", DbType.Date, DateTime.Today.ToShortDateString()));
                    sdsContractManagerDetail.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                    sdsContractManagerDetail.Update();
                    gridContractManager.DataBind();
                    fvContractManager.DataBind();
                    fvContractManager.ChangeMode(FormViewMode.ReadOnly);
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
        }

        protected void eWorkDep_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eWorkDep_Edit = (TextBox)fvContractManager.FindControl("eWorkDep_Edit");
            if ((eWorkDep_Edit != null) && (eWorkDep_Edit.Text.Trim() != ""))
            {
                Label eWorkDepName_Edit = (Label)fvContractManager.FindControl("eWorkDepName_Edit");
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vWorkDep = eWorkDep_Edit.Text.Trim();
                string vTempStr = "select [Name] from Department where DepNo = '" + vWorkDep + "' ";
                string vWorkDepName = PF.GetValue(vConnStr, vTempStr, "Name");
                if (vWorkDepName == "")
                {
                    vWorkDepName = vWorkDep.Trim();
                    vTempStr = "select DepNo from Department where [Name] = '" + vWorkDepName + "' ";
                    vWorkDep = PF.GetValue(vConnStr, vTempStr, "DepNo");
                }
                eWorkDep_Edit.Text = vWorkDep.Trim();
                eWorkDepName_Edit.Text = vWorkDepName.Trim();
            }
        }

        protected void eResponsibleDep_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eResponsibleDep_Edit = (TextBox)fvContractManager.FindControl("eResponsibleDep_Edit");
            if ((eResponsibleDep_Edit != null) && (eResponsibleDep_Edit.Text.Trim() != ""))
            {
                Label eResponsibleDepName_Edit = (Label)fvContractManager.FindControl("eResponsibleDepName_Edit");
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vResponsibleDep = eResponsibleDep_Edit.Text.Trim();
                string vTempStr = "select [Name] from Department where DepNo = '" + vResponsibleDep + "' ";
                string vResponsibleDepName = PF.GetValue(vConnStr, vTempStr, "Name");
                if (vResponsibleDepName == "")
                {
                    vResponsibleDepName = vResponsibleDep.Trim();
                    vTempStr = "select DepNo from Department where [Name] = '" + vResponsibleDepName + "' ";
                    vResponsibleDep = PF.GetValue(vConnStr, vTempStr, "DepNo");
                }
                eResponsibleDep_Edit.Text = vResponsibleDep.Trim();
                eResponsibleDepName_Edit.Text = vResponsibleDepName.Trim();
            }
        }

        protected void cbIsClose_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsClose_Edit = (CheckBox)fvContractManager.FindControl("cbIsClose_Edit");
            if (cbIsClose_Edit != null)
            {
                Label eIsClose_Edit = (Label)fvContractManager.FindControl("eIsClose_Edit");
                eIsClose_Edit.Text = (cbIsClose_Edit.Checked) ? "V" : "X";
            }
        }

        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            Label eIndexNo_INS = (Label)fvContractManager.FindControl("eIndexNo_INS");
            if (eIndexNo_INS != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vIndexNo = "";
                TextBox eContractNo_INS = (TextBox)fvContractManager.FindControl("eContractNo_INS");
                string vContractNo = eContractNo_INS.Text.Trim();
                TextBox eWorkDep_INS = (TextBox)fvContractManager.FindControl("eWorkDep_INS");
                string vWorkDep = eWorkDep_INS.Text.Trim();
                TextBox eResponsibleDep_INS = (TextBox)fvContractManager.FindControl("eResponsibleDep_INS");
                string vResponsibleDep = eResponsibleDep_INS.Text.Trim();
                TextBox eInvoiceNo_INS = (TextBox)fvContractManager.FindControl("eInvoiceNo_INS");
                string vInvoiceNo = eInvoiceNo_INS.Text.Trim();
                TextBox eCustomName_INS = (TextBox)fvContractManager.FindControl("eCustomName_INS");
                string vCustomName = eCustomName_INS.Text.Trim();
                TextBox eContractText_INS = (TextBox)fvContractManager.FindControl("eContractText_INS");
                string vContractText = eContractText_INS.Text.Trim();
                TextBox eAssignDate_INS = (TextBox)fvContractManager.FindControl("eAssignDate_INS");
                string vAssignDateStr = eAssignDate_INS.Text.Trim();
                DateTime vAssignDate;
                if (!DateTime.TryParse(vAssignDateStr, out vAssignDate))
                {
                    vAssignDate = DateTime.Today;
                }
                TextBox eEffectiveDateS_INS = (TextBox)fvContractManager.FindControl("eEffectiveDateS_INS");
                string vEffectiveDateSStr = eEffectiveDateS_INS.Text.Trim();
                DateTime vEffectiveDateS;
                if (!DateTime.TryParse(vEffectiveDateSStr, out vEffectiveDateS))
                {
                    vEffectiveDateS = DateTime.Today;
                }
                TextBox eEffectiveDateE_INS = (TextBox)fvContractManager.FindControl("eEffectiveDateE_INS");
                string vEffectiveDateEStr = eEffectiveDateE_INS.Text.Trim();
                DateTime vEffectiveDateE;
                if (!DateTime.TryParse(vEffectiveDateEStr, out vEffectiveDateE))
                {
                    vEffectiveDateE = DateTime.Today;
                }
                Label eIsClose_INS = (Label)fvContractManager.FindControl("eIsClose_INS");
                string vIsClose = eIsClose_INS.Text.Trim();
                TextBox eAmount_INS = (TextBox)fvContractManager.FindControl("eAmount_INS");
                string vAmount = eAmount_INS.Text.Trim();
                TextBox eTax_INS = (TextBox)fvContractManager.FindControl("eTax_INS");
                string vTax = eTax_INS.Text.Trim();
                Label eTotalAmount_INS = (Label)fvContractManager.FindControl("eTotalAmount_INS");
                string vTotalAmount = eTotalAmount_INS.Text.Trim();
                TextBox eStampDuty_INS = (TextBox)fvContractManager.FindControl("eStampDuty_INS");
                string vStampDuty = eStampDuty_INS.Text.Trim();
                TextBox eRemark_INS = (TextBox)fvContractManager.FindControl("eRemark_INS");
                string vRemark = eRemark_INS.Text.Trim();

                string vAssignYM = vAssignDate.Year.ToString() + vAssignDate.Month.ToString("D2");
                string vTempStr = "select max(IndexNo) MaxIndex from ContractManager where IndexNo like '" + vAssignYM + "%' ";
                string vOldIndexNo = PF.GetValue(vConnStr, vTempStr, "MaxIndex");
                int vNewIndex = (vOldIndexNo != "") ? Int32.Parse((vOldIndexNo.Replace(vAssignYM + "CM", ""))) + 1 : 1;
                vIndexNo = vAssignYM + "CM" + vNewIndex.ToString("D4");
                try
                {
                    sdsContractManagerDetail.InsertCommand = "INSERT INTO ContractManager " + Environment.NewLine +
                                                             "       (IndexNo, ContractNo, WorkDep, ResponsibleDep, InvoiceNo, CustomName, ContractText, " + Environment.NewLine +
                                                             "        AssignDate, EffectiveDateS, EffectiveDateE, IsClose, " + Environment.NewLine +
                                                             "        Amount, Tax, TotalAmount, StampDuty, Remark, BuMan, BuDate) " + Environment.NewLine +
                                                             "VALUES (@IndexNo, @ContractNo, @WorkDep, @ResponsibleDep, @InvoiceNo, @CustomName, @ContractText, " + Environment.NewLine +
                                                             "        @AssignDate, @EffectiveDateS, @EffectiveDateE, @IsClose, " + Environment.NewLine +
                                                             "        @Amount, @Tax, @TotalAmount, @StampDuty, @Remark, @BuMan, @BuDate)";
                    sdsContractManagerDetail.InsertParameters.Clear();
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("ContractNo", DbType.String, (vContractNo != "") ? vContractNo : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("WorkDep", DbType.String, (vWorkDep != "") ? vWorkDep : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("ResponsibleDep", DbType.String, (vResponsibleDep != "") ? vResponsibleDep : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("InvoiceNo", DbType.String, (vInvoiceNo != "") ? vInvoiceNo : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("CustomName", DbType.String, (vCustomName != "") ? vCustomName : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("ContractText", DbType.String, (vContractText != "") ? vContractText : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("AssignDate", DbType.Date, (vAssignDateStr != "") ? vAssignDate.ToShortDateString() : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("EffectiveDateS", DbType.Date, (vEffectiveDateSStr != "") ? vEffectiveDateS.ToShortDateString() : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("EffectiveDateE", DbType.Date, (vEffectiveDateEStr != "") ? vEffectiveDateE.ToShortDateString() : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("IsClose", DbType.String, (vIsClose != "") ? vIsClose : "X"));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("Amount", DbType.Int32, (vAmount != "") ? vAmount : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("Tax", DbType.Int32, (vTax != "") ? vTax : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("TotalAmount", DbType.Int32, (vTotalAmount != "") ? vTotalAmount : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("StampDuty", DbType.Int32, (vStampDuty != "") ? vStampDuty : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("Remark", DbType.String, (vRemark != "") ? vRemark : String.Empty));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    sdsContractManagerDetail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, DateTime.Today.ToShortDateString()));
                    sdsContractManagerDetail.Insert();
                    gridContractManager.DataBind();
                    fvContractManager.DataBind();
                    fvContractManager.ChangeMode(FormViewMode.ReadOnly);
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
        }

        protected void cbIsClose_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsClose_INS = (CheckBox)fvContractManager.FindControl("cbIsClose_INS");
            if (cbIsClose_INS != null)
            {
                Label eIsClose_INS = (Label)fvContractManager.FindControl("eIsClose_INS");
                eIsClose_INS.Text = (cbIsClose_INS.Checked) ? "V" : "X";
            }
        }

        protected void eWorkDep_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eWorkDep_INS = (TextBox)fvContractManager.FindControl("eWorkDep_INS");
            if ((eWorkDep_INS != null) && (eWorkDep_INS.Text.Trim() != ""))
            {
                Label eWorkDepName_INS = (Label)fvContractManager.FindControl("eWorkDepName_INS");
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vWorkDep = eWorkDep_INS.Text.Trim();
                string vTempStr = "select [Name] from Department where DepNo = '" + vWorkDep + "' ";
                string vWorkDepName = PF.GetValue(vConnStr, vTempStr, "Name");
                if (vWorkDepName == "")
                {
                    vWorkDepName = vWorkDep.Trim();
                    vTempStr = "select DepNo from Department where [Name] = '" + vWorkDepName + "' ";
                    vWorkDep = PF.GetValue(vConnStr, vTempStr, "DepNo");
                }
                eWorkDep_INS.Text = vWorkDep.Trim();
                eWorkDepName_INS.Text = vWorkDepName.Trim();
            }
        }

        protected void eResponsibleDep_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eResponsibleDep_INS = (TextBox)fvContractManager.FindControl("eResponsibleDep_INS");
            if ((eResponsibleDep_INS != null) && (eResponsibleDep_INS.Text.Trim() != ""))
            {
                Label eResponsibleDepName_INS = (Label)fvContractManager.FindControl("eResponsibleDepName_INS");
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vResponsibleDep = eResponsibleDep_INS.Text.Trim();
                string vTempStr = "select [Name] from Department where DepNo = '" + vResponsibleDep + "' ";
                string vResponsibleDepName = PF.GetValue(vConnStr, vTempStr, "Name");
                if (vResponsibleDepName == "")
                {
                    vResponsibleDepName = vResponsibleDep.Trim();
                    vTempStr = "select DepNo from Department where [Name] = '" + vResponsibleDepName + "' ";
                    vResponsibleDep = PF.GetValue(vConnStr, vTempStr, "DepNo");
                }
                eResponsibleDep_INS.Text = vResponsibleDep.Trim();
                eResponsibleDepName_INS.Text = vResponsibleDepName.Trim();
            }
        }

        protected void fvContractManager_DataBound(object sender, EventArgs e)
        {
            string vCaseDateURL = "";
            string vCaseDateScript = "";
            switch (fvContractManager.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Label eIndexNo_List = (Label)fvContractManager.FindControl("eIndexNo_List");
                    if (eIndexNo_List != null)
                    {
                        Label eIsClose_List = (Label)fvContractManager.FindControl("eIsClose_List");
                        CheckBox cbIsClose_List = (CheckBox)fvContractManager.FindControl("cbIsClose_List");
                        cbIsClose_List.Checked = (eIsClose_List.Text.Trim() == "V");
                        Button bbCaseClose_List = (Button)fvContractManager.FindControl("bbCaseClose_List");
                        bbCaseClose_List.Visible = (eIsClose_List.Text.Trim() != "V");
                        Button bbNew_List = (Button)fvContractManager.FindControl("bbNew_List");
                        bbNew_List.Visible = ((vLoginDepNo == "03") || (vLoginDepNo == "09"));
                        Button bbEdit_List = (Button)fvContractManager.FindControl("bbEdit_List");
                        bbEdit_List.Visible = ((vLoginDepNo == "03") || (vLoginDepNo == "09"));
                        Button bbDelete_List = (Button)fvContractManager.FindControl("bbDelete_List");
                        bbDelete_List.Visible = false;
                    }
                    Button bbNew_Empty = (Button)fvContractManager.FindControl("bbNew_Empty");
                    if (bbNew_Empty != null)
                    {
                        bbNew_Empty.Visible = ((vLoginDepNo == "03") || (vLoginDepNo == "09"));
                    }
                    break;
                case FormViewMode.Edit:
                    Label eIndexNo_Edit = (Label)fvContractManager.FindControl("eIndexNo_Edit");
                    if (eIndexNo_Edit != null)
                    {
                        TextBox eAssignDate_Edit = (TextBox)fvContractManager.FindControl("eAssignDate_Edit");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eAssignDate_Edit.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eAssignDate_Edit.Attributes["onCLick"] = vCaseDateScript;
                        TextBox eEffectiveDateS_Edit = (TextBox)fvContractManager.FindControl("eEffectiveDateS_Edit");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eEffectiveDateS_Edit.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eEffectiveDateS_Edit.Attributes["onCLick"] = vCaseDateScript;
                        TextBox eEffectiveDateE_Edit = (TextBox)fvContractManager.FindControl("eEffectiveDateE_Edit");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eEffectiveDateE_Edit.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eEffectiveDateE_Edit.Attributes["onCLick"] = vCaseDateScript;
                        Label eIsClose_Edit = (Label)fvContractManager.FindControl("eIsClose_Edit");
                        CheckBox cbIsClose_Edit = (CheckBox)fvContractManager.FindControl("cbIsClose_Edit");
                        cbIsClose_Edit.Checked = (eIsClose_Edit.Text.Trim() == "V");
                    }
                    break;
                case FormViewMode.Insert:
                    Label eIndexNo_INS = (Label)fvContractManager.FindControl("eIndexNo_INS");
                    if (eIndexNo_INS != null)
                    {
                        TextBox eAssignDate_INS = (TextBox)fvContractManager.FindControl("eAssignDate_INS");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eAssignDate_INS.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eAssignDate_INS.Attributes["onCLick"] = vCaseDateScript;
                        eAssignDate_INS.Text = DateTime.Today.ToShortDateString();
                        TextBox eEffectiveDateS_INS = (TextBox)fvContractManager.FindControl("eEffectiveDateS_INS");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eEffectiveDateS_INS.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eEffectiveDateS_INS.Attributes["onCLick"] = vCaseDateScript;
                        TextBox eEffectiveDateE_INS = (TextBox)fvContractManager.FindControl("eEffectiveDateE_INS");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eEffectiveDateE_INS.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eEffectiveDateE_INS.Attributes["onCLick"] = vCaseDateScript;
                    }
                    break;
            }
        }

        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            Label eIndexNo_List = (Label)fvContractManager.FindControl("eIndexNo_List");
            if (eIndexNo_List != null)
            {
                //先複製一份到異動記錄
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    string vIndexNo_List = eIndexNo_List.Text.Trim();
                    string vTempStr = "select max(Items) MaxItem from ContractManager_History where IndexNo = '" + vIndexNo_List + "' ";
                    string vMaxItemsStr = PF.GetValue(vConnStr, vTempStr, "MaxItem");
                    int vMaxItem = (vMaxItemsStr != "") ? Int32.Parse(vMaxItemsStr) : 0;
                    string vNewItems = (vMaxItem + 1).ToString("D4");
                    string vIndexNoItem = vIndexNo_List.Trim() + vNewItems.Trim();
                    vTempStr = "INSERT INTO ContractManager_History " + Environment.NewLine +
                               "       (IndexNoItem, Items, ModifyMode, IndexNo, ContractNo, WorkDep, ResponsibleDep, InvoiceNo, CustomName, ContractText, " + Environment.NewLine +
                               "        AssignDate, EffectiveDateS, EffectiveDateE, IsClose, Amount, Tax, TotalAmount, StampDuty, Remark, BuMan, BuDate) " + Environment.NewLine +
                               "SELECT '" + vIndexNoItem + "', '" + vNewItems + "', 'DEL', IndexNo, ContractNo, WorkDep, ResponsibleDep, InvoiceNo, " + Environment.NewLine +
                               "       CustomName, ContractText, AssignDate, EffectiveDateS, EffectiveDateE, IsClose, Amount, Tax, TotalAmount, " + Environment.NewLine +
                               "       StampDuty, Remark, '" + vLoginID + "', GetDate() " + Environment.NewLine +
                               "  FROM ContractManager " + Environment.NewLine +
                               " WHERE IndexNo = '" + vIndexNo_List + "' ";
                    PF.ExecSQL(vConnStr, vTempStr);
                    sdsContractManagerDetail.DeleteParameters.Clear();
                    sdsContractManagerDetail.DeleteParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo_List));
                    sdsContractManagerDetail.Delete();
                    gridContractManager.DataBind();
                    fvContractManager.DataBind();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
        }

        protected void bbCaseClose_List_Click(object sender, EventArgs e)
        {
            Label eIndexNo_List = (Label)fvContractManager.FindControl("eIndexNo_List");
            if (eIndexNo_List != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    //先複製一份到異動記錄
                    string vIndexNo_List = eIndexNo_List.Text.Trim();
                    string vTempStr = "select max(Items) MaxItem from ContractManager_History where IndexNo = '" + vIndexNo_List + "' ";
                    string vMaxItemsStr = PF.GetValue(vConnStr, vTempStr, "MaxItem");
                    int vMaxItem = (vMaxItemsStr != "") ? Int32.Parse(vMaxItemsStr) : 0;
                    string vNewItems = (vMaxItem + 1).ToString("D4");
                    string vIndexNoItem = vIndexNo_List.Trim() + vNewItems.Trim();
                    vTempStr = "INSERT INTO ContractManager_History " + Environment.NewLine +
                               "       (IndexNoItem, Items, ModifyMode, IndexNo, ContractNo, WorkDep, ResponsibleDep, InvoiceNo, CustomName, ContractText, " + Environment.NewLine +
                               "        AssignDate, EffectiveDateS, EffectiveDateE, IsClose, Amount, Tax, TotalAmount, StampDuty, Remark, BuMan, BuDate) " + Environment.NewLine +
                               "SELECT '" + vIndexNoItem + "', '" + vNewItems + "', 'CLO', IndexNo, ContractNo, WorkDep, ResponsibleDep, InvoiceNo, " + Environment.NewLine +
                               "       CustomName, ContractText, AssignDate, EffectiveDateS, EffectiveDateE, IsClose, Amount, Tax, TotalAmount, " + Environment.NewLine +
                               "       StampDuty, Remark, '" + vLoginID + "', GetDate() " + Environment.NewLine +
                               "  FROM ContractManager " + Environment.NewLine +
                               " WHERE IndexNo = '" + vIndexNo_List + "' ";
                    PF.ExecSQL(vConnStr, vTempStr);
                    //實際進行結案
                    vTempStr = "update ContractManager " + Environment.NewLine +
                               "   set IsClose = 'V' " + Environment.NewLine +
                               " where IndexNｏ = '" + vIndexNo_List + "' ";
                    PF.ExecSQL(vConnStr, vTempStr);
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
        }

        /// <summary>
        /// 編輯模式下計算金額小計
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eAmount_Edit_TextChanged(object sender, EventArgs e)
        {
            Label eIndexNo_Edit = (Label)fvContractManager.FindControl("eIndexNo_Edit");
            if (eIndexNo_Edit != null)
            {
                TextBox eAmount_Edit = (TextBox)fvContractManager.FindControl("eAmount_Edit");
                TextBox eTax_Edit = (TextBox)fvContractManager.FindControl("eTax_Edit");
                Label eTotalAmount_Edit = (Label)fvContractManager.FindControl("eTotalAmount_Edit");
                int vAmount_Edit = 0;
                int vTax_Edit = 0;
                if (Int32.TryParse(eAmount_Edit.Text.Trim(), out vAmount_Edit))
                {
                    if (Int32.TryParse(eTax_Edit.Text.Trim(), out vTax_Edit))
                    {
                        eTotalAmount_Edit.Text = (vAmount_Edit + vTax_Edit).ToString();
                    }
                    else
                    {
                        eTotalAmount_Edit.Text = vAmount_Edit.ToString();
                    }
                }
                else
                {
                    if (Int32.TryParse(eTax_Edit.Text.Trim(), out vTax_Edit))
                    {
                        eTotalAmount_Edit.Text = eTax_Edit.ToString();
                    }
                    else
                    {
                        eTotalAmount_Edit.Text = "";
                    }
                }
            }
        }

        /// <summary>
        /// 新增模式下計算金額小計
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eAmount_INS_TextChanged(object sender, EventArgs e)
        {
            Label eIndexNo_INS = (Label)fvContractManager.FindControl("eIndexNo_INS");
            if (eIndexNo_INS != null)
            {
                TextBox eAmount_INS = (TextBox)fvContractManager.FindControl("eAmount_INS");
                TextBox eTax_INS = (TextBox)fvContractManager.FindControl("eTax_INS");
                Label eTotalAmount_INS = (Label)fvContractManager.FindControl("eTotalAmount_INS");
                int vAmount_INS = 0;
                int vTax_INS = 0;
                if (Int32.TryParse(eAmount_INS.Text.Trim(), out vAmount_INS))
                {
                    if (Int32.TryParse(eTax_INS.Text.Trim(), out vTax_INS))
                    {
                        eTotalAmount_INS.Text = (vAmount_INS + vTax_INS).ToString();
                    }
                    else
                    {
                        eTotalAmount_INS.Text = vAmount_INS.ToString();
                    }
                }
                else
                {
                    if (Int32.TryParse(eTax_INS.Text.Trim(), out vTax_INS))
                    {
                        eTotalAmount_INS.Text = eTax_INS.ToString();
                    }
                    else
                    {
                        eTotalAmount_INS.Text = "";
                    }
                }
            }
        }
    }
}