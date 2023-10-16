using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverRegularTraining : System.Web.UI.Page
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
                if (!IsPostBack)
                {
                    fuImportData.Visible = (vLoginDepNo == "09");
                    bbImportData.Visible = (vLoginDepNo == "09");
                    plPrint.Visible = false;
                    plShowData.Visible = true;
                    if ((vLoginDepNo != "06") && (vLoginDepNo != "09") && (vLoginDepNo != "02") && (vLoginDepNo != "05"))
                    {
                        eDepName_Start_Search.Text = vLoginDepNo;
                        eDepName_Start_Search.Enabled = false;
                        eDepName_End_Search.Enabled = false;
                    }
                    else
                    {
                        eDepName_Start_Search.Enabled = true;
                        eDepName_End_Search.Enabled = true;
                    }
                }
                else
                {
                    GetDataList();
                }
            }
            else
            {
                Response.Redirect("~/default.aspx");
            }
        }

        private string GetSelStr()
        {
            string vResultStr = "";
            string vWStr_DepNo = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? "   and e.DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? "   and e.DepNo = '" + eDepNo_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? "   and e.DepNo = '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eEmpNo_Search.Text.Trim() != "") ? "   and e.EmpNo = '" + eEmpNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_SearchMode = "";
            switch (rbSearchMode.SelectedIndex)
            {
                default:
                    vWStr_SearchMode = "";
                    break;
                case 0:
                    vWStr_SearchMode = "";
                    break;
                case 1:
                    //vWStr_SearchMode = "   and ((e.DriverLicense < GetDate()) or (e.BBCall < GetDate())) " + Environment.NewLine;
                    vWStr_SearchMode = "   and ((e.LicenceCheck < GetDate()) or (e.BBCall < GetDate())) " + Environment.NewLine;
                    break;
                case 2:
                    //vWStr_SearchMode = "   and ((DATEDIFF(day, getdate(), DriverLicense) between 0 and 30) or (DATEDIFF(day, getdate(), BBCall) between 0 and 30)) " + Environment.NewLine;
                    vWStr_SearchMode = "   and ((DATEDIFF(day, getdate(), LicenceCheck) between 0 and 30) or (DATEDIFF(day, getdate(), BBCall) between 0 and 30)) " + Environment.NewLine;
                    break;
            }

            vResultStr = "select DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName, " + Environment.NewLine +
                         "       EmpNo, [Name], " + Environment.NewLine +
                         "       LicenceType, (select ClassTxt from DBDICB where FKey = '人事資料檔      Employee        LicenceType' and ClassNo = e.LicenceType) as LicenceType_C, " + Environment.NewLine +
                         //"       LicencingDay, DriverLicense, DATEADD(day, -15, DriverLicense) DriverLicenceLimitDate, " + Environment.NewLine +
                         "       LicencingDay, LicenceCheck, DATEADD(day, -15, LicenceCheck) DriverLicenceLimitDate, " + Environment.NewLine +
                         "       TraTool, (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        TRATOOL' and ClassNo = e.TraTool) as TraTool_C, " + Environment.NewLine +
                         "       AutoNo TrainingGetDate, BBCall TrainingStopDate, DATEADD(day, -15, BBCall) TrainingLimitDate " + Environment.NewLine +
                         "  from Employee e " + Environment.NewLine +
                         " where [Type] = '20' " + Environment.NewLine +
                         "   and isnull(e.IsOfficialDriver, 'X') <> 'V' " + Environment.NewLine +
                         "   and e.LeaveDay is null" + Environment.NewLine + vWStr_DepNo + vWStr_EmpNo + vWStr_SearchMode +
                         " order by e.DepNo, e.EmpNo";
            return vResultStr;
        }

        private void GetDataList()
        {
            string vSelectStr = GetSelStr();
            sdsDataList.SelectCommand = "";
            sdsDataList.SelectCommand = vSelectStr;
            gridDataList.DataBind();
        }

        protected void eDepNo_Start_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Search = eDepNo_Start_Search.Text.Trim();
            string vDepName_Search = "";
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Search + "' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vDepName_Search = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Search == "")
            {
                vDepName_Search = vDepNo_Search;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Search + "' ";
                vDepNo_Search = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_Start_Search.Text = vDepNo_Search;
            eDepName_Start_Search.Text = vDepName_Search;
        }

        protected void eDepNo_End_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Search = eDepNo_End_Search.Text.Trim();
            string vDepName_Search = "";
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Search + "' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vDepName_Search = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Search == "")
            {
                vDepName_Search = vDepNo_Search;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Search + "' ";
                vDepNo_Search = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_End_Search.Text = vDepNo_Search;
            eDepName_End_Search.Text = vDepName_Search;
        }

        protected void eEmpNo_Search_TextChanged(object sender, EventArgs e)
        {
            string vEmpNo = eEmpNo_Search.Text.Trim();
            string vEmpName = "";
            string vSQLStr = "select [Name] from EMployee where EmpNo = '" + vEmpNo + "' and LeaveDay is null";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vEmpName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vEmpName + "' ";
                vEmpNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eEmpNo_Search.Text = vEmpNo;
            eEmpName_Search.Text = vEmpName;
        }

        protected void bbImportData_Click(object sender, EventArgs e)
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            string vUploadFileName = "";
            string vExtName = "";
            string vDepNo_Temp = "";
            string vEmpNo = "";
            string vLicenceType = "";
            string vLicenceType_C = "";
            DateTime vLicencingDay;
            string vLicencingDayStr = "";
            DateTime vDriverLicense;
            string vDriverLicenseStr = "";
            string vTraTool = "";
            string vTraTool_C = "";
            DateTime vTrainingGetDate;
            string vTrainingGetDateStr = "";
            DateTime vTrainingStopDate;
            string vTrainingStopDateStr = "";
            string vTempData = "";
            string vTempStr = "";
            string vTempDateStr = "";
            string vUpdateStr = "";

            if (fuImportData.FileName != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                vExtName = Path.GetExtension(fuImportData.FileName);
                vUploadFileName = vUploadPath + fuImportData.FileName;
                fuImportData.SaveAs(vUploadFileName);
                if (vExtName == ".xls")
                {
                    HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuImportData.FileContent);
                    for (int iSheetIndex = 0; iSheetIndex < 11; iSheetIndex++)
                    {
                        HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(iSheetIndex);
                        for (int i = sheetExcel_H.FirstRowNum + 1; i <= sheetExcel_H.LastRowNum; i++)
                        {
                            HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                            if (vRowTemp_H.Cells.Count >= 11)
                            {
                                try
                                {
                                    if (vRowTemp_H.GetCell(0).CellType != CellType.Blank)
                                    {
                                        vTempData = vRowTemp_H.GetCell(0).ToString();
                                        vDepNo_Temp = Int32.Parse(vTempData).ToString("D2");
                                    }
                                    if (vRowTemp_H.GetCell(2).CellType != CellType.Blank)
                                    {
                                        vTempData = vRowTemp_H.GetCell(2).ToString();
                                        vEmpNo = Int32.Parse(vTempData).ToString("D6");
                                    }
                                    if (vRowTemp_H.GetCell(6).CellType != CellType.Blank)
                                    {
                                        vLicenceType_C = vRowTemp_H.GetCell(6).StringCellValue;
                                        vTempStr = "select ClassNo from DBDICB where FKey = '人事資料檔      Employee        LicenceType' and ClassTxt = '" + vLicenceType_C + "' ";
                                        vLicenceType = PF.GetValue(vConnStr, vTempStr, "ClassNo").Trim();
                                    }
                                    if (vRowTemp_H.GetCell(7).CellType != CellType.Blank)
                                    {
                                        vTempDateStr = vRowTemp_H.GetCell(7).ToString().Trim();
                                        vLicencingDay = DateTime.Parse(PF.GetAD(vTempDateStr));
                                        vLicencingDayStr = vLicencingDay.ToString("yyyy/MM/dd");
                                    }
                                    if (vRowTemp_H.GetCell(8).CellType != CellType.Blank)
                                    {
                                        vTempDateStr = vRowTemp_H.GetCell(8).ToString().Trim();
                                        vDriverLicense = DateTime.Parse(PF.GetAD(vTempDateStr));
                                        vDriverLicenseStr = vDriverLicense.ToString("yyyy/MM/dd");
                                    }
                                    if (vRowTemp_H.GetCell(9).CellType != CellType.Blank)
                                    {
                                        vTraTool_C = vRowTemp_H.GetCell(9).StringCellValue;
                                        vTempStr = "select ClassNo from DBDICB where FKey = '人事資料檔      EMPLOYEE        TRATOOL' and ClassTxt = '" + vTraTool_C + "' ";
                                        vTraTool = PF.GetValue(vConnStr, vTempStr, "ClassNo").Trim();
                                    }
                                    if (vRowTemp_H.GetCell(10).CellType != CellType.Blank)
                                    {
                                        vTempDateStr = vRowTemp_H.GetCell(10).ToString().Trim();
                                        vTrainingGetDate = DateTime.Parse(PF.GetAD(vTempDateStr));
                                        vTrainingGetDateStr = vTrainingGetDate.ToString("yyyy/MM/dd");
                                    }
                                    if (vRowTemp_H.GetCell(11).CellType != CellType.Blank)
                                    {
                                        vTempDateStr = vRowTemp_H.GetCell(11).ToString().Trim();
                                        vTrainingStopDate = DateTime.Parse(PF.GetAD(vTempDateStr));
                                        vTrainingStopDateStr = vTrainingStopDate.ToString("yyyy/MM/dd");
                                    }
                                }
                                catch (Exception ex_GetData)
                                {
                                    Response.Write("<Script language='Javascript'>");
                                    Response.Write("alert('" + vDepNo_Temp + "  " + vEmpNo + "  " + vTempDateStr + Environment.NewLine + ex_GetData.Message + ex_GetData.ToString() + "')");
                                    Response.Write("</" + "Script>");
                                    //throw;
                                }
                                try
                                {
                                    vUpdateStr = "update Employee " + Environment.NewLine +
                                                 "   set LicenceType = '" + vLicenceType + "', " + Environment.NewLine +
                                                 "       LicencingDay = '" + vLicencingDayStr + "', " + Environment.NewLine +
                                                 //"       DriverLicense = '" + vDriverLicenseStr + "', " + Environment.NewLine +
                                                 "       LicenceCheck = '" + vDriverLicenseStr + "', " + Environment.NewLine +
                                                 "       TraTool = '" + vTraTool + "', " + Environment.NewLine +
                                                 "       AutoNo = '" + vTrainingGetDateStr + "', " + Environment.NewLine +
                                                 "       BBCall = '" + vTrainingStopDateStr + "' " + Environment.NewLine +
                                                 " where EmpNo = '" + vEmpNo + "' ";
                                    PF.ExecSQL(vConnStr, vUpdateStr);
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
                else if (vExtName == ".xlsx")
                {
                    XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuImportData.FileContent);
                    for (int iSheetIndex = 0; iSheetIndex < 11; iSheetIndex++)
                    {
                        XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(iSheetIndex);
                        for (int i = sheetExcel_X.FirstRowNum + 1; i <= sheetExcel_X.LastRowNum; i++)
                        {
                            XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                            if (vRowTemp_X.Cells.Count >= 11)
                            {
                                if (vRowTemp_X.GetCell(0).CellType != CellType.Blank)
                                {
                                    vTempData = vRowTemp_X.GetCell(0).ToString();
                                    vDepNo_Temp = Int32.Parse(vTempData).ToString("D2");
                                }
                                if (vRowTemp_X.GetCell(2).CellType != CellType.Blank)
                                {
                                    vTempData = vRowTemp_X.GetCell(2).ToString();
                                    vEmpNo = Int32.Parse(vTempData).ToString("D6");
                                }
                                if (vRowTemp_X.GetCell(6).CellType != CellType.Blank)
                                {
                                    vLicenceType_C = vRowTemp_X.GetCell(6).StringCellValue;
                                    vTempStr = "select ClassNo from DBDICB where FKey = '人事資料檔      Employee        LicenceType' and ClassTxt = '" + vLicenceType_C + "' ";
                                    vLicenceType = PF.GetValue(vConnStr, vTempStr, "ClassNo").Trim();
                                }
                                if (vRowTemp_X.GetCell(7).CellType != CellType.Blank)
                                {
                                    vTempDateStr = vRowTemp_X.GetCell(7).ToString().Trim();
                                    vLicencingDay = DateTime.Parse(PF.GetAD(vTempDateStr));
                                    vLicencingDayStr = vLicencingDay.ToString("yyyy/MM/dd");
                                }
                                if (vRowTemp_X.GetCell(8).CellType != CellType.Blank)
                                {
                                    vTempDateStr = vRowTemp_X.GetCell(8).ToString().Trim();
                                    vDriverLicense = DateTime.Parse(PF.GetAD(vTempDateStr));
                                    vDriverLicenseStr = vDriverLicense.ToString("yyyy/MM/dd");
                                }
                                if (vRowTemp_X.GetCell(9).CellType != CellType.Blank)
                                {
                                    vTraTool_C = vRowTemp_X.GetCell(9).StringCellValue;
                                    vTempStr = "select ClassNo from DBDICB where FKey = '人事資料檔      EMPLOYEE        TRATOOL' and ClassTxt = '" + vTraTool_C + "' ";
                                    vTraTool = PF.GetValue(vConnStr, vTempStr, "ClassNo").Trim();
                                }
                                if (vRowTemp_X.GetCell(10).CellType != CellType.Blank)
                                {
                                    vTempDateStr = vRowTemp_X.GetCell(10).ToString().Trim();
                                    vTrainingGetDate = DateTime.Parse(PF.GetAD(vTempDateStr));
                                    vTrainingGetDateStr = vTrainingGetDate.ToString("yyyy/MM/dd");
                                }
                                if (vRowTemp_X.GetCell(11).CellType != CellType.Blank)
                                {
                                    vTempDateStr = vRowTemp_X.GetCell(11).ToString().Trim();
                                    vTrainingStopDate = DateTime.Parse(PF.GetAD(vTempDateStr));
                                    vTrainingStopDateStr = vTrainingStopDate.ToString("yyyy/MM/dd");
                                }
                                try
                                {
                                    vUpdateStr = "update Employee " + Environment.NewLine +
                                                 "   set LicenceType = '" + vLicenceType + "', " + Environment.NewLine +
                                                 "       LicencingDay = '" + vLicencingDayStr + "', " + Environment.NewLine +
                                                 //"       DriverLicense = '" + vDriverLicenseStr + "', " + Environment.NewLine +
                                                 "       LicenceCheck = '" + vDriverLicenseStr + "', " + Environment.NewLine +
                                                 "       TraTool = '" + vTraTool + "', " + Environment.NewLine +
                                                 "       AutoNo = '" + vTrainingGetDateStr + "', " + Environment.NewLine +
                                                 "       BBCall = '" + vTrainingStopDateStr + "' " + Environment.NewLine +
                                                 " where EmpNo = '" + vEmpNo + "' ";
                                    PF.ExecSQL(vConnStr, vUpdateStr);
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
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            GetDataList();
        }

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

            //
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

            string vSelectStr = GetSelStr();
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
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "DEPNO") ? "部門編號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "部門" :
                                      (drExcel.GetName(i).ToUpper() == "EMPNO") ? "員工編號" :
                                      (drExcel.GetName(i).ToUpper() == "NAME") ? "員工姓名" :
                                      (drExcel.GetName(i).ToUpper() == "LICENCETYPE") ? "證照種類代碼" :
                                      (drExcel.GetName(i).ToUpper() == "LICENCETYPE_C") ? "證照種類" :
                                      (drExcel.GetName(i).ToUpper() == "LICENCINGDAY") ? "駕照發照日" :
                                      //(drExcel.GetName(i) == "DriverLicense") ? "駕照審驗日" :
                                      (drExcel.GetName(i).ToUpper() == "LICENCECHECK") ? "駕照審驗日" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVERLICENCELIMITDATE") ? "駕照審驗期限" :
                                      (drExcel.GetName(i).ToUpper() == "TRATOOL") ? "定期訓練證類別代碼" :
                                      (drExcel.GetName(i).ToUpper() == "TRATOOL_C") ? "定期訓練證類別" :
                                      (drExcel.GetName(i).ToUpper() == "TRAININGGETDATE") ? "定期訓練發照日" :
                                      (drExcel.GetName(i).ToUpper() == "TRAININGSTOPDATE") ? "定期訓練審驗日" :
                                      (drExcel.GetName(i).ToUpper() == "TRAININGLIMITDATE") ? "定期訓練審驗期限" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i).ToUpper() == "LICENCINGDAY") ||
                                 //(drExcel.GetName(i) == "DriverLicense") ||
                                 (drExcel.GetName(i).ToUpper() == "LICENCECHECK") ||
                                 (drExcel.GetName(i).ToUpper() == "DRIVERLICENCELIMITDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "TRAININGGETDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "TRAININGSTOPDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "TRAININGLIMITDATE")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(((vBuDate.Year) - 1911).ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
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
                            string vDepNoStr = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? eDepNo_Start_Search.Text.Trim() + " 站到 " + eDepNo_End_Search.Text.Trim() + " 站" :
                                               ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? eDepNo_Start_Search.Text.Trim() + " 站" :
                                               ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? eDepNo_End_Search.Text.Trim() + " 站" : "全部";
                            string vEmpNoStr = (eEmpNo_Search.Text.Trim() != "") ? eEmpNo_Search.Text.Trim() : "全部";
                            string vSearchModeStr = rbSearchMode.SelectedItem.Text.Trim();
                            string vRecordNote = "匯出檔案_駕駛員駕照及定期訓練證管理" + Environment.NewLine +
                                                 "DriverRegularTraining.aspx" + Environment.NewLine +
                                                 "過濾類別：" + vSearchModeStr + Environment.NewLine +
                                                 "駕駛員編號：" + vEmpNoStr + Environment.NewLine +
                                                 "站別：" + vDepNoStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            string vFileName = "駕駛員駕照及定期訓練證管理";
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

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                string vSelectStr = GetSelStr();
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable();
                daPrint.Fill(dtPrint);

                ReportDataSource rdsPrint = new ReportDataSource("DriverRegularTrainingP", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\DriverRegularTrainingP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plShowData.Visible = false;

                string vDepNoStr = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? eDepNo_Start_Search.Text.Trim() + " 站到 " + eDepNo_End_Search.Text.Trim() + " 站" :
                                   ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? eDepNo_Start_Search.Text.Trim() + " 站" :
                                   ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? eDepNo_End_Search.Text.Trim() + " 站" : "全部";
                string vEmpNoStr = (eEmpNo_Search.Text.Trim() != "") ? eEmpNo_Search.Text.Trim() : "全部";
                string vSearchModeStr = rbSearchMode.SelectedItem.Text.Trim();
                string vRecordNote = "預覽報表_駕駛員駕照及定期訓練證管理" + Environment.NewLine +
                                     "DriverRegularTraining.aspx" + Environment.NewLine +
                                     "過濾類別：" + vSearchModeStr + Environment.NewLine +
                                     "駕駛員編號：" + vEmpNoStr + Environment.NewLine +
                                     "站別：" + vDepNoStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plShowData.Visible = true;
        }
    }
}