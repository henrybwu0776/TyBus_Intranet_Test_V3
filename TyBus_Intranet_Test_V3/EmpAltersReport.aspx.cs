using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class EmpAltersReport : System.Web.UI.Page
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
                    if (!IsPostBack)
                    {

                    }
                    string vInureDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eInureDate_S_Search.ClientID;
                    string vInureDateScript = "window.open('" + vInureDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eInureDate_S_Search.Attributes["onClick"] = vInureDateScript;

                    vInureDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eInureDate_E_Search.ClientID;
                    vInureDateScript = "window.open('" + vInureDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eInureDate_E_Search.Attributes["onClick"] = vInureDateScript;
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

        protected void eEmpNo_Search_TextChanged(object sender, EventArgs e)
        {
            string vEmpNo_Temp = eEmpNo_Search.Text.Trim();
            string vEmpName_Temp = "";
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp.Trim() + "' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vEmpName_Temp.Trim() == "")
            {
                vEmpName_Temp = vEmpNo_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_Temp + "' order by EmpNo DESC";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eEmpNo_Search.Text = vEmpNo_Temp.Trim();
            eEmpName_Search.Text = vEmpName_Temp.Trim();
        }

        protected void eDepNo_S_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_S_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp.Trim() == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '%" + vDepName_Temp.Trim() + "%' order by DepNo";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_S_Search.Text = vDepNo_Temp.Trim();
            eDepName_S_Search.Text = vDepName_Temp.Trim();
        }

        protected void eDepNo_E_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_E_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp.Trim() == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '%" + vDepName_Temp.Trim() + "%' order by DepNo";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_E_Search.Text = vDepNo_Temp.Trim();
            eDepName_E_Search.Text = vDepName_Temp.Trim();
        }

        private string GetWhereStr()
        {
            string vResultStr = "";
            string vWStr_AlterNo = ((eAltersNo_S_Search.Text.Trim() != "") && (eAltersNo_E_Search.Text.Trim() != "")) ? "   and a.AlterNo between '" + eAltersNo_S_Search.Text.Trim() + "' and '" + eAltersNo_E_Search.Text.Trim() + Environment.NewLine :
                                   ((eAltersNo_S_Search.Text.Trim() != "") && (eAltersNo_E_Search.Text.Trim() == "")) ? "   and a.AlterNo like '" + eAltersNo_S_Search.Text.Trim() + "%' " + Environment.NewLine :
                                   ((eAltersNo_S_Search.Text.Trim() == "") && (eAltersNo_E_Search.Text.Trim() != "")) ? "   and a.AlterNo like '" + eAltersNo_E_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_DepNo = ((eDepName_S_Search.Text.Trim() != "") && (eDepName_E_Search.Text.Trim() != "")) ? "   and a.DepNo between '" + eDepName_S_Search.Text.Trim() + "' and '" + eDepName_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepName_S_Search.Text.Trim() != "") && (eDepName_E_Search.Text.Trim() == "")) ? "   and a.DepNo like '" + eDepName_S_Search.Text.Trim() + "%' " + Environment.NewLine :
                                 ((eDepName_S_Search.Text.Trim() == "") && (eDepName_E_Search.Text.Trim() != "")) ? "   and a.DepNo like '" + eDepName_E_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eEmpNo_Search.Text.Trim() != "") ? "   and a.EmpNo like '" + eEmpNo_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vDate_S = (eInureDate_S_Search.Text.Trim() != "") ? PF.GetAD(eInureDate_S_Search.Text.Trim()) : "";
            string vDate_E = (eInureDate_E_Search.Text.Trim() != "") ? PF.GetAD(eInureDate_E_Search.Text.Trim()) : "";
            string vWStr_InureDate = ((eInureDate_S_Search.Text.Trim() != "") && (eInureDate_E_Search.Text.Trim() != "")) ? "   and a.InureDate between '" + vDate_S.Trim() + "' and '" + vDate_E.Trim() + "' " + Environment.NewLine :
                                     ((eInureDate_S_Search.Text.Trim() != "") && (eInureDate_E_Search.Text.Trim() == "")) ? "   and a.InureDate = '" + vDate_S.Trim() + "' " + Environment.NewLine :
                                     ((eInureDate_S_Search.Text.Trim() == "") && (eInureDate_E_Search.Text.Trim() != "")) ? "   and a.InureDate = '" + vDate_E.Trim() + "' " + Environment.NewLine : "";
            string vAlterTypeStr = "";
            foreach (ListItem vitems in eAlterType_Search.Items)
            {
                if (vitems.Selected)
                {
                    if (string.IsNullOrEmpty(vAlterTypeStr))
                    {
                        vAlterTypeStr += "'" + vitems.Value + "' ";
                    }
                    else
                    {
                        vAlterTypeStr += ",'" + vitems.Value + "' ";
                    }
                }
            }
            string vWStr_AlterType = (vAlterTypeStr.Trim() != "") ? "   and a.AlterType in (" + vAlterTypeStr + ") " + Environment.NewLine : "";

            vResultStr = " where isnull(a.AlterNo, '') <> '' " + Environment.NewLine +
                         vWStr_AlterNo +
                         vWStr_AlterType +
                         vWStr_DepNo +
                         vWStr_EmpNo +
                         vWStr_InureDate;
            return vResultStr;
        }

        protected void bbOK_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_Temp = "select [Name] from Custom where Types = 'O' and Code = 'A000' ";
            string vWhereStr = GetWhereStr();
            string vCompanyName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            string vReportTitle = vCompanyName + ddlReportName_Search.SelectedItem.Text.Trim();
            //決定要匯出哪個報表格式
            switch (ddlReportName_Search.SelectedValue.Trim())
            {
                case "A000":
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('請先選擇格式')");
                    Response.Write("</" + "Script>");
                    break;
                case "A001":
                    ExportExcel_A001(vWhereStr, vReportTitle, ddlReportName_Search.SelectedItem.Text.Trim());
                    break;
                case "A002":
                    ExportExcel_A002(vWhereStr, vReportTitle, ddlReportName_Search.SelectedItem.Text.Trim());
                    break;
                case "A003":
                    ExportExcel_A003(vWhereStr, vReportTitle, ddlReportName_Search.SelectedItem.Text.Trim());
                    break;
                case "A004":
                    ExportExcel_A004(vWhereStr, vReportTitle, ddlReportName_Search.SelectedItem.Text.Trim());
                    break;
                case "A005":
                    vReportTitle = vCompanyName + ddlReportName_Search.Items[4].Text.Trim();
                    ExportExcel_A005(vWhereStr, vReportTitle, ddlReportName_Search.SelectedItem.Text.Trim());
                    break;
                case "A006":
                    ExportExcel_A001(vWhereStr, vReportTitle, ddlReportName_Search.SelectedItem.Text.Trim()); //跟第一張內容完全一樣，只有標題不一樣
                    break;
                case "A007":
                    ExportExcel_A004(vWhereStr, vReportTitle, ddlReportName_Search.SelectedItem.Text.Trim()); //跟第四張內容一樣。也是只有標題不同
                    break;
                case "A008":
                    vReportTitle = vCompanyName + ddlReportName_Search.Items[4].Text.Trim();
                    ExportExcel_A005(vWhereStr, vReportTitle, ddlReportName_Search.SelectedItem.Text.Trim()); //跟第五張內容一樣。也是只有標題不同
                    break;
                case "A009":
                    ExportExcel_A005(vWhereStr, vReportTitle, ddlReportName_Search.SelectedItem.Text.Trim()); //跟第五張內容一樣。也是只有標題不同
                    break;
            }

            string vRecordAlterNoStr = ((eAltersNo_S_Search.Text.Trim() != "") && (eAltersNo_E_Search.Text.Trim() != "")) ? eAltersNo_S_Search.Text.Trim() + "~" + eAltersNo_E_Search.Text.Trim() :
                                       ((eAltersNo_S_Search.Text.Trim() != "") && (eAltersNo_E_Search.Text.Trim() == "")) ? eAltersNo_S_Search.Text.Trim() :
                                       ((eAltersNo_S_Search.Text.Trim() == "") && (eAltersNo_E_Search.Text.Trim() != "")) ? eAltersNo_E_Search.Text.Trim() : "全部";
            string vRecordDepNoStr = ((eDepName_S_Search.Text.Trim() != "") && (eDepName_E_Search.Text.Trim() != "")) ? eDepName_S_Search.Text.Trim() + "~" + eDepName_E_Search.Text.Trim() :
                                     ((eDepName_S_Search.Text.Trim() != "") && (eDepName_E_Search.Text.Trim() == "")) ? eDepName_S_Search.Text.Trim() :
                                     ((eDepName_S_Search.Text.Trim() == "") && (eDepName_E_Search.Text.Trim() != "")) ? eDepNo_E_Search.Text.Trim() : "全部";
            string vRecordEmpNoStr = (eEmpNo_Search.Text.Trim() != "") ? eEmpNo_Search.Text.Trim() : "";
            string vDate_S = (eInureDate_S_Search.Text.Trim() != "") ? PF.GetAD(eInureDate_S_Search.Text.Trim()) : "";
            string vDate_E = (eInureDate_E_Search.Text.Trim() != "") ? PF.GetAD(eInureDate_E_Search.Text.Trim()) : "";
            string vRecordDateStr = ((eInureDate_S_Search.Text.Trim() != "") && (eInureDate_E_Search.Text.Trim() != "")) ? vDate_S.Trim() + "~" + vDate_E.Trim() :
                                    ((eInureDate_S_Search.Text.Trim() != "") && (eInureDate_E_Search.Text.Trim() == "")) ? vDate_S.Trim() :
                                    ((eInureDate_S_Search.Text.Trim() == "") && (eInureDate_E_Search.Text.Trim() != "")) ? vDate_E.Trim() : "";
            string vRecordAlterTypeStr = "";
            foreach (ListItem vitems in eAlterType_Search.Items)
            {
                if (vitems.Selected)
                {
                    if (string.IsNullOrEmpty(vRecordAlterTypeStr))
                    {
                        vRecordAlterTypeStr += vitems.Value;
                    }
                    else
                    {
                        vRecordAlterTypeStr += "," + vitems.Value;
                    }
                }
            }

            string vRecordNote = "匯出檔案_" + ddlReportName_Search.SelectedItem.Text.Trim() + Environment.NewLine +
                                 "EmpAltersReport.aspx" + Environment.NewLine +
                                 "人事異動單號：" + vRecordAlterNoStr + Environment.NewLine +
                                 "部門：" + vRecordDepNoStr + Environment.NewLine +
                                 "駕駛員：" + vRecordEmpNoStr + Environment.NewLine +
                                 "發生日期：" + vRecordDateStr + Environment.NewLine +
                                 "異動類別：" + vRecordAlterTypeStr;
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
        }

        private void ExportExcel_A001(string fWhereStr, string fReportTitle, string fFileName)
        {
            string vSelectStr = "select a.AlterWord, a.DecDate, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where ClassNo = a.AlterType and fkey = '人事異動檔      ALTERS          altertype') as AlterType_C, " + Environment.NewLine +
                                "       a.EmpNo, a.[Name], (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        type' and ClassNo = a.Type) as Type_C, " + Environment.NewLine +
                                "       (select [Name] from Department where DepNo = a.DepNo) as DepName, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        TITLE' and ClassNo = a.Title) as Title_C, " + Environment.NewLine +
                                "       e.EduLevel1 as EduLevel, a.LevelNo2, a.SalaryPoint, a.InureDate, e.AssumeDay " + Environment.NewLine +
                                "  from Alters as a left join Employee as e on e.EmpNo = a.EmpNo " + Environment.NewLine +
                                fWhereStr +
                                " order by a.AlterWord, a.AlterNo ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DateTime vTempDate;
            string vTempStr = "";
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

            HSSFCellStyle csTitle_R = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle_R.VerticalAlignment = VerticalAlignment.Center;
            csTitle_R.Alignment = HorizontalAlignment.Right; //水平靠右
            csTitle_R.SetFont(fontTitle);

            //設定資料內容欄位的格式
            HSSFCellStyle csData = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            HSSFFont fontData = (HSSFFont)wbExcel.CreateFont();
            fontData.FontHeightInPoints = 12;
            csData.SetFont(fontData);

            HSSFCellStyle csData_R = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_R.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csData_R.Alignment = HorizontalAlignment.Right; //水平靠右
            csData_R.SetFont(fontData);

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
            string vAlterWord = "";
            string vDecDate = "";
            string vPageName = "";
            int vLinesNo = 0;
            int vColNo = 0;
            int vPageNo = 1;

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daExcel = new SqlDataAdapter(vSelectStr, connExcel);
                connExcel.Open();
                DataTable dtExcel = new DataTable();
                daExcel.Fill(dtExcel);
                for (int i = 0; i < dtExcel.Rows.Count; i++)
                {
                    if ((dtExcel.Rows[i]["AlterWord"].ToString().Trim() != vAlterWord) || (vLinesNo == 8)) //人令字號不同或是已經寫入五筆資料時時
                    {
                        if (vLinesNo > 0) //列數大於 0 表示前面還有一頁，所以在這一頁要放頁尾
                        {
                            wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                            if (vLinesNo < 8)
                            {
                                for (int ik = vLinesNo; ik < 8; ik++)
                                {
                                    wsExcel.CreateRow(ik);
                                }
                                vLinesNo = 8;
                            }
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("上開送審人事是否有當");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 0, 4));
                            wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("副");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("業");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("主");
                            wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("辦");
                            wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("核示");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, 4));
                            wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("課");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("長");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                        }
                        vDecDate = PF.TransDateString(dtExcel.Rows[i]["DecDate"].ToString().Trim(), "C2");
                        vLinesNo = 0; //列數先歸零
                        if (dtExcel.Rows[i]["AlterWord"].ToString().Trim() == vAlterWord)
                        {
                            vPageNo++;
                            vPageName = vAlterWord + "_" + vPageNo.ToString();
                        }
                        else
                        {
                            vAlterWord = dtExcel.Rows[i]["AlterWord"].ToString().Trim();
                            vPageName = vAlterWord;
                            vPageNo = 1;
                        }
                        //wsExcel = (HSSFSheet)wbExcel.CreateSheet(vAlterWord); //根據不同的人令字號產生不同的工作表
                        wsExcel = (HSSFSheet)wbExcel.CreateSheet(vPageName); //根據不同的人令字號產生不同的工作表
                        //產生標頭欄位
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(fReportTitle);
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 0, 6));
                        wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("桃汽客人字第 " + vAlterWord + " 號");
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 7, 10));
                        wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue(vDecDate);
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 7, 10));
                        wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int ci = 0; ci < dtExcel.Columns.Count; ci++)
                        {
                            switch (dtExcel.Columns[ci].ColumnName.Trim())
                            {
                                default:
                                    vHeaderText = "";
                                    vColNo = -1;
                                    break;
                                case "AlterType_C":
                                    vHeaderText = "區分";
                                    vColNo = 0;
                                    break;
                                case "EmpNo":
                                    vHeaderText = "員工編號";
                                    vColNo = 1;
                                    break;
                                case "Name":
                                    vHeaderText = "姓名";
                                    vColNo = 2;
                                    break;
                                case "Type_C":
                                    vHeaderText = "人員類別";
                                    vColNo = 3;
                                    break;
                                case "DepName":
                                    vHeaderText = "單位";
                                    vColNo = 4;
                                    break;
                                case "Title_C":
                                    vHeaderText = "職別";
                                    vColNo = 5;
                                    break;
                                case "EduLevel":
                                    vHeaderText = "主要學歷";
                                    vColNo = 6;
                                    break;
                                case "LevelNo2":
                                    vHeaderText = "職等";
                                    vColNo = 7;
                                    break;
                                case "SalaryPoint":
                                    vHeaderText = "俸點";
                                    vColNo = 8;
                                    break;
                                case "InureDate":
                                    vHeaderText = "生效日期";
                                    vColNo = 9;
                                    break;
                                case "AssumeDay":
                                    vHeaderText = "到職日期";
                                    vColNo = 10;
                                    break;
                            }
                            if (vColNo != -1)
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(vColNo).SetCellValue(vHeaderText);
                                wsExcel.GetRow(vLinesNo).GetCell(vColNo).CellStyle = csData;
                            }
                        }
                        vLinesNo++;
                    }
                    //開始寫入資料
                    wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                    wsExcel.CreateRow(vLinesNo);
                    for (int ci = 0; ci < dtExcel.Columns.Count; ci++)
                    {
                        ///*
                        switch (dtExcel.Columns[ci].ColumnName.Trim())
                        {
                            default:
                                vColNo = -1;
                                break;
                            case "AlterType_C":
                                vTempStr = dtExcel.Rows[i]["AlterType_C"].ToString().Trim();
                                vColNo = 0;
                                break;
                            case "EmpNo":
                                vTempStr = dtExcel.Rows[i]["EmpNo"].ToString().Trim();
                                vColNo = 1;
                                break;
                            case "Name":
                                vTempStr = dtExcel.Rows[i]["Name"].ToString().Trim();
                                vColNo = 2;
                                break;
                            case "Type_C":
                                vTempStr = dtExcel.Rows[i]["Type_C"].ToString().Trim();
                                vColNo = 3;
                                break;
                            case "DepName":
                                vTempStr = dtExcel.Rows[i]["DepName"].ToString().Trim();
                                vColNo = 4;
                                break;
                            case "Title_C":
                                vTempStr = dtExcel.Rows[i]["Title_C"].ToString().Trim();
                                vColNo = 5;
                                break;
                            case "EduLevel":
                                vTempStr = dtExcel.Rows[i]["EduLevel"].ToString().Trim();
                                vColNo = 6;
                                break;
                            case "LevelNo2":
                                vTempStr = dtExcel.Rows[i]["LevelNo2"].ToString().Trim();
                                vColNo = 7;
                                break;
                            case "SalaryPoint":
                                vTempStr = dtExcel.Rows[i]["SalaryPoint"].ToString().Trim();
                                vColNo = 8;
                                break;
                            case "InureDate":
                                vTempDate = DateTime.Parse(dtExcel.Rows[i]["InureDate"].ToString().Trim());
                                vTempStr = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                                vColNo = 9;
                                break;
                            case "AssumeDay":
                                vTempDate = DateTime.Parse(dtExcel.Rows[i]["AssumeDay"].ToString().Trim());
                                vTempStr = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                                vColNo = 10;
                                break;
                        }
                        if (vColNo != -1)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(vColNo);
                            wsExcel.GetRow(vLinesNo).GetCell(vColNo).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(vColNo).SetCellValue(vTempStr);
                            wsExcel.GetRow(vLinesNo).GetCell(vColNo).CellStyle = csData;
                        }
                    }
                    vLinesNo++;
                }
                if (dtExcel.Rows.Count > 0)
                {
                    vLinesNo++;
                    wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                    if (vLinesNo < 8)
                    {
                        for (int ik = vLinesNo + 1; ik < 8; ik++)
                        {
                            wsExcel.CreateRow(ik);
                        }
                        vLinesNo = 8;
                    }
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("上開送審人事是否有當");
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 0, 4));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("副");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("業");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("主");
                    wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("辦");
                    wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("核示");
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, 4));
                    wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("課");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("長");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
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
                        Response.Write("alert('" + eMessage.Message + "')");
                        Response.Write("</" + "Script>");
                        //throw;
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('無符合條件資料')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        private void ExportExcel_A002(string fWhereStr, string fReportTitle, string fFileName)
        {
            string vSelectStr = "select a.AlterWord, a.DecDate, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where ClassNo = a.AlterType and fkey = '人事異動檔      ALTERS          altertype') as AlterType_C, " + Environment.NewLine +
                                "       a.EmpNo, a.[Name], (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        type' and ClassNo = a.Type) as Type_C, " + Environment.NewLine +
                                "       (select [Name] from Department where DepNo = a.DepNo) as DepName, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        TITLE' and ClassNo = a.Title) as Title_C, " + Environment.NewLine +
                                "       e.EduLevel1 as EduLevel, a.ALRemark8, a.InureDate, e.AssumeDay " + Environment.NewLine +
                                "  from Alters as a left join Employee as e on e.EmpNo = a.EmpNo " + Environment.NewLine +
                                fWhereStr +
                                " order by a.AlterWord, a.AlterNo ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DateTime vTempDate;
            string vTempStr = "";
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

            HSSFCellStyle csTitle_R = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle_R.VerticalAlignment = VerticalAlignment.Center;
            csTitle_R.Alignment = HorizontalAlignment.Right; //水平靠右
            csTitle_R.SetFont(fontTitle);

            //設定資料內容欄位的格式
            HSSFCellStyle csData = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            HSSFFont fontData = (HSSFFont)wbExcel.CreateFont();
            fontData.FontHeightInPoints = 12;
            csData.SetFont(fontData);

            HSSFCellStyle csData_R = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_R.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csData_R.Alignment = HorizontalAlignment.Right; //水平靠右
            csData_R.SetFont(fontData);

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
            string vAlterWord = "";
            string vDecDate = "";
            string vPageName = "";
            int vLinesNo = 0;
            int vColNo = 0;
            int vPageNo = 1;

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daExcel = new SqlDataAdapter(vSelectStr, connExcel);
                connExcel.Open();
                DataTable dtExcel = new DataTable();
                daExcel.Fill(dtExcel);
                for (int i = 0; i < dtExcel.Rows.Count; i++)
                {
                    if ((dtExcel.Rows[i]["AlterWord"].ToString().Trim() != vAlterWord) || (vLinesNo == 8)) //人令字號不同或是已經寫入五筆資料時時
                    {
                        if (vLinesNo > 0) //列數大於 0 表示前面還有一頁，所以在這一頁要放頁尾
                        {
                            wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                            if (vLinesNo < 8)
                            {
                                for (int ik = vLinesNo; ik < 8; ik++)
                                {
                                    wsExcel.CreateRow(ik);
                                }
                                vLinesNo = 8;
                            }
                            //vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("上開送審人事是否有當");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 4));
                            wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("副");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("業");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("主");
                            wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("辦");
                            wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("核示");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, 4));
                            wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("課");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("長");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                        }
                        vDecDate = PF.TransDateString(dtExcel.Rows[i]["DecDate"].ToString().Trim(), "C2");
                        vLinesNo = 0; //列數先歸零
                        if (dtExcel.Rows[i]["AlterWord"].ToString().Trim() == vAlterWord)
                        {
                            vPageNo++;
                            vPageName = vAlterWord + "_" + vPageNo.ToString();
                        }
                        else
                        {
                            vAlterWord = dtExcel.Rows[i]["AlterWord"].ToString().Trim();
                            vPageName = vAlterWord;
                            vPageNo = 1;
                        }
                        //wsExcel = (HSSFSheet)wbExcel.CreateSheet(vAlterWord); //根據不同的人令字號產生不同的工作表
                        wsExcel = (HSSFSheet)wbExcel.CreateSheet(vPageName); //根據不同的人令字號產生不同的工作表
                        //產生標頭欄位
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(fReportTitle);
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 0, 6));
                        wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("桃汽客人字第 " + vAlterWord + " 號");
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 7, 9));
                        wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue(vDecDate);
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 7, 9));
                        wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int ci = 0; ci < dtExcel.Columns.Count; ci++)
                        {
                            switch (dtExcel.Columns[ci].ColumnName.Trim())
                            {
                                default:
                                    vHeaderText = "";
                                    vColNo = -1;
                                    break;
                                case "AlterType_C":
                                    vHeaderText = "區分";
                                    vColNo = 0;
                                    break;
                                case "EmpNo":
                                    vHeaderText = "員工編號";
                                    vColNo = 1;
                                    break;
                                case "Name":
                                    vHeaderText = "姓名";
                                    vColNo = 2;
                                    break;
                                case "Type_C":
                                    vHeaderText = "人員類別";
                                    vColNo = 3;
                                    break;
                                case "DepName":
                                    vHeaderText = "單位";
                                    vColNo = 4;
                                    break;
                                case "Title_C":
                                    vHeaderText = "職別";
                                    vColNo = 5;
                                    break;
                                case "EduLevel":
                                    vHeaderText = "主要學歷";
                                    vColNo = 6;
                                    break;
                                case "ALRemark8":
                                    vHeaderText = "審核意見";
                                    vColNo = 7;
                                    break;
                                case "InureDate":
                                    vHeaderText = "生效日期";
                                    vColNo = 8;
                                    break;
                                case "AssumeDay":
                                    vHeaderText = "到職日期";
                                    vColNo = 9;
                                    break;
                            }
                            if (vColNo != -1)
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(vColNo).SetCellValue(vHeaderText);
                                wsExcel.GetRow(vLinesNo).GetCell(vColNo).CellStyle = csData;
                            }
                        }
                        vLinesNo++;
                    }
                    //開始寫入資料
                    wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                    wsExcel.CreateRow(vLinesNo);
                    for (int ci = 0; ci < dtExcel.Columns.Count; ci++)
                    {
                        ///*
                        switch (dtExcel.Columns[ci].ColumnName.Trim())
                        {
                            default:
                                vColNo = -1;
                                break;
                            case "AlterType_C":
                                vTempStr = dtExcel.Rows[i]["AlterType_C"].ToString().Trim();
                                vColNo = 0;
                                break;
                            case "EmpNo":
                                vTempStr = dtExcel.Rows[i]["EmpNo"].ToString().Trim();
                                vColNo = 1;
                                break;
                            case "Name":
                                vTempStr = dtExcel.Rows[i]["Name"].ToString().Trim();
                                vColNo = 2;
                                break;
                            case "Type_C":
                                vTempStr = dtExcel.Rows[i]["Type_C"].ToString().Trim();
                                vColNo = 3;
                                break;
                            case "DepName":
                                vTempStr = dtExcel.Rows[i]["DepName"].ToString().Trim();
                                vColNo = 4;
                                break;
                            case "Title_C":
                                vTempStr = dtExcel.Rows[i]["Title_C"].ToString().Trim();
                                vColNo = 5;
                                break;
                            case "EduLevel":
                                vTempStr = dtExcel.Rows[i]["EduLevel"].ToString().Trim();
                                vColNo = 6;
                                break;
                            case "ALRemark8":
                                vTempStr = dtExcel.Rows[i]["ALRemark8"].ToString().Trim();
                                vColNo = 7;
                                break;
                            case "InureDate":
                                vTempDate = DateTime.Parse(dtExcel.Rows[i]["InureDate"].ToString().Trim());
                                vTempStr = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                                vColNo = 8;
                                break;
                            case "AssumeDay":
                                vTempDate = DateTime.Parse(dtExcel.Rows[i]["AssumeDay"].ToString().Trim());
                                vTempStr = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                                vColNo = 9;
                                break;
                        }
                        if (vColNo != -1)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(vColNo);
                            wsExcel.GetRow(vLinesNo).GetCell(vColNo).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(vColNo).SetCellValue(vTempStr);
                            wsExcel.GetRow(vLinesNo).GetCell(vColNo).CellStyle = csData;
                        }
                    }
                    vLinesNo++;
                }
                if (dtExcel.Rows.Count > 0)
                {
                    vLinesNo++;
                    wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                    if (vLinesNo < 8)
                    {
                        for (int ik = vLinesNo + 1; ik < 8; ik++)
                        {
                            wsExcel.CreateRow(ik);
                        }
                        vLinesNo = 8;
                    }
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("上開送審人事是否有當");
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 0, 4));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("副");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("業");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("主");
                    wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("辦");
                    wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("核示");
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, 4));
                    wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("課");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("長");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
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
                        Response.Write("alert('" + eMessage.Message + "')");
                        Response.Write("</" + "Script>");
                        //throw;
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('無符合條件資料')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        private void ExportExcel_A003(string fWhereStr, string fReportTitle, string fFileName)
        {
            string vSelectStr = "select a.AlterWord, a.DecDate, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where ClassNo = a.AlterType and fkey = '人事異動檔      ALTERS          altertype') as AlterType_C, " + Environment.NewLine +
                                "       a.EmpNo, a.[Name], (select [Name] from Department where DepNo = a.DepNo) as NewDepName, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        type' and ClassNo = a.Type) as Type_C, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        SERVNO' and ClassNo = a.ServNo) as NewServNo_C, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        TITLE' and ClassNo = a.Title) as NewTitle_C, " + Environment.NewLine +
                                "       (select [Name] from Department where DepNo = a.DepNoO) as OriDepName, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        type' and ClassNo = a.TypeO) as OriType_C, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        SERVNO' and ClassNo = a.ServNoO) as OriServNo_C, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        TITLE' and ClassNo = (a.TitleO + '0')) as OriTitle_C, " + Environment.NewLine +
                                "       a.ALRemark7, a.LevelNo2, a.SalaryPoint, a.InureDate " + Environment.NewLine +
                                "  from Alters as a " + Environment.NewLine +
                                fWhereStr +
                                " order by a.AlterWord, a.AlterNo ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DateTime vTempDate;
            string vTempStr = "";
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

            HSSFCellStyle csTitle_R = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle_R.VerticalAlignment = VerticalAlignment.Center;
            csTitle_R.Alignment = HorizontalAlignment.Right; //水平靠右
            csTitle_R.SetFont(fontTitle);

            //設定資料內容欄位的格式
            HSSFCellStyle csData = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            HSSFFont fontData = (HSSFFont)wbExcel.CreateFont();
            fontData.FontHeightInPoints = 12;
            csData.SetFont(fontData);

            HSSFCellStyle csData_R = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_R.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csData_R.Alignment = HorizontalAlignment.Right; //水平靠右
            csData_R.SetFont(fontData);

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
            string vAlterWord = "";
            string vDecDate = "";
            string vPageName = "";
            int vLinesNo = 0;
            int vColNo = 0;
            int vPageNo = 1;

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daExcel = new SqlDataAdapter(vSelectStr, connExcel);
                connExcel.Open();
                DataTable dtExcel = new DataTable();
                daExcel.Fill(dtExcel);
                for (int i = 0; i < dtExcel.Rows.Count; i++)
                {
                    if ((dtExcel.Rows[i]["AlterWord"].ToString().Trim() != vAlterWord) || (vLinesNo == 13)) //人令字號不同或是已經寫入五筆資料時時
                    {
                        if (vLinesNo > 0) //列數大於 0 表示前面還有一頁，所以在這一頁要放頁尾
                        {
                            wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                            if (vLinesNo < 13)
                            {
                                for (int ik = vLinesNo; ik < 8; ik++)
                                {
                                    wsExcel.CreateRow(ik);
                                }
                                vLinesNo = 13;
                            }
                            //vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("上開送審人事是否有當");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 4));
                            wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("副");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("業");
                            wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(11).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(12).SetCellValue("主");
                            wsExcel.GetRow(vLinesNo).GetCell(12).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(11).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(12).SetCellValue("辦");
                            wsExcel.GetRow(vLinesNo).GetCell(12).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("核示");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, 4));
                            wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellValue("課");
                            wsExcel.GetRow(vLinesNo).GetCell(11).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellValue("長");
                            wsExcel.GetRow(vLinesNo).GetCell(11).CellStyle = csData_R;
                        }
                        vDecDate = PF.TransDateString(dtExcel.Rows[i]["DecDate"].ToString().Trim(), "C2");
                        vLinesNo = 0; //列數先歸零
                        if (dtExcel.Rows[i]["AlterWord"].ToString().Trim() == vAlterWord)
                        {
                            vPageNo++;
                            vPageName = vAlterWord + "_" + vPageNo.ToString();
                        }
                        else
                        {
                            vAlterWord = dtExcel.Rows[i]["AlterWord"].ToString().Trim();
                            vPageName = vAlterWord;
                            vPageNo = 1;
                        }
                        //wsExcel = (HSSFSheet)wbExcel.CreateSheet(vAlterWord); //根據不同的人令字號產生不同的工作表
                        wsExcel = (HSSFSheet)wbExcel.CreateSheet(vPageName); //根據不同的人令字號產生不同的工作表
                        //產生標頭欄位
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(fReportTitle);
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 0, 6));
                        wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("桃汽客人字第 " + vAlterWord + " 號");
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 7, 9));
                        wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue(vDecDate);
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 7, 9));
                        wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int ci = 0; ci < dtExcel.Columns.Count; ci++)
                        {
                            switch (dtExcel.Columns[ci].ColumnName.Trim())
                            {
                                default:
                                    vHeaderText = "";
                                    vColNo = -1;
                                    break;
                                case "AlterType_C":
                                    vHeaderText = "區分";
                                    vColNo = 0;
                                    break;
                                case "EmpNo":
                                    vHeaderText = "員工編號";
                                    vColNo = 1;
                                    break;
                                case "Name":
                                    vHeaderText = "姓名";
                                    vColNo = 2;
                                    break;
                                case "NewDepName":
                                    vHeaderText = "新部門";
                                    vColNo = 3;
                                    break;
                                case "Type_C":
                                    vHeaderText = "人員類別";
                                    vColNo = 4;
                                    break;
                                case "NewTitle_C":
                                    vHeaderText = "新職稱";
                                    vColNo = 5;
                                    break;
                                case "OriDepName":
                                    vHeaderText = "原部門";
                                    vColNo = 6;
                                    break;
                                case "OriType_C":
                                    vHeaderText = "原人員類別";
                                    vColNo = 7;
                                    break;
                                case "OriTitle_C":
                                    vHeaderText = "原職稱";
                                    vColNo = 8;
                                    break;
                                case "ALRemark7":
                                    vHeaderText = "勤務內容";
                                    vColNo = 9;
                                    break;
                                case "LevelNo2":
                                    vHeaderText = "職等";
                                    vColNo = 10;
                                    break;
                                case "SalaryPoint":
                                    vHeaderText = "俸點";
                                    vColNo = 11;
                                    break;
                                case "InureDate":
                                    vHeaderText = "生效日期";
                                    vColNo = 12;
                                    break;
                            }
                            if (vColNo != -1)
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(vColNo).SetCellValue(vHeaderText);
                                wsExcel.GetRow(vLinesNo).GetCell(vColNo).CellStyle = csData;
                            }
                        }
                        vLinesNo++;
                    }
                    //開始寫入資料
                    wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.CreateRow(vLinesNo + 1);
                    for (int ci = 0; ci < dtExcel.Columns.Count; ci++)
                    {
                        ///*
                        switch (dtExcel.Columns[ci].ColumnName.Trim())
                        {
                            default:
                                vColNo = -1;
                                break;
                            case "AlterType_C":
                                vTempStr = dtExcel.Rows[i]["AlterType_C"].ToString().Trim();
                                vColNo = 0;
                                break;
                            case "EmpNo":
                                vTempStr = dtExcel.Rows[i]["EmpNo"].ToString().Trim();
                                vColNo = 1;
                                break;
                            case "Name":
                                vTempStr = dtExcel.Rows[i]["Name"].ToString().Trim();
                                vColNo = 2;
                                break;
                            case "NewDepName":
                                vTempStr = dtExcel.Rows[i]["NewDepName"].ToString().Trim();
                                vColNo = 3;
                                break;
                            case "Type_C":
                                vTempStr = dtExcel.Rows[i]["Type_C"].ToString().Trim();
                                vColNo = 4;
                                break;
                            case "NewServNo_C":
                                vTempStr = dtExcel.Rows[i]["NewServNo_C"].ToString().Trim();
                                vColNo = 5;
                                break;
                            case "NewTitle_C":
                                vTempStr = dtExcel.Rows[i]["NewTitle_C"].ToString().Trim();
                                vColNo = 5;
                                break;
                            case "OriDepName":
                                vTempStr = dtExcel.Rows[i]["OriDepName"].ToString().Trim();
                                vColNo = 6;
                                break;
                            case "OriType_C":
                                vTempStr = dtExcel.Rows[i]["OriType_C"].ToString().Trim();
                                vColNo = 7;
                                break;
                            case "OriServNo_C":
                                vTempStr = dtExcel.Rows[i]["OriServNo_C"].ToString().Trim();
                                vColNo = 8;
                                break;
                            case "OriTitle_C":
                                vTempStr = dtExcel.Rows[i]["OriTitle_C"].ToString().Trim();
                                vColNo = 8;
                                break;
                            case "ALRemark7":
                                vTempStr = dtExcel.Rows[i]["ALRemark7"].ToString().Trim();
                                vColNo = 9;
                                break;
                            case "LevelNo2":
                                vTempStr = dtExcel.Rows[i]["LevelNo2"].ToString().Trim();
                                vColNo = 10;
                                break;
                            case "SalaryPoint":
                                vTempStr = dtExcel.Rows[i]["SalaryPoint"].ToString().Trim();
                                vColNo = 11;
                                break;
                            case "InureDate":
                                vTempDate = DateTime.Parse(dtExcel.Rows[i]["InureDate"].ToString().Trim());
                                vTempStr = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                                vColNo = 12;
                                break;
                        }
                        if (vColNo != -1)
                        {
                            if ((dtExcel.Columns[ci].ColumnName.Trim() == "NewTitle_C") ||
                                (dtExcel.Columns[ci].ColumnName.Trim() == "OriTitle_C"))
                            {
                                wsExcel.GetRow(vLinesNo + 1).CreateCell(vColNo);
                                wsExcel.GetRow(vLinesNo + 1).GetCell(vColNo).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo + 1).GetCell(vColNo).SetCellValue(vTempStr);
                                wsExcel.GetRow(vLinesNo + 1).GetCell(vColNo).CellStyle = csData;
                            }
                            else if ((dtExcel.Columns[ci].ColumnName.Trim() == "NewServNo_C") ||
                                     (dtExcel.Columns[ci].ColumnName.Trim() == "OriServNo_C"))
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(vColNo);
                                wsExcel.GetRow(vLinesNo).GetCell(vColNo).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(vColNo).SetCellValue(vTempStr);
                                wsExcel.GetRow(vLinesNo).GetCell(vColNo).CellStyle = csData;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(vColNo);
                                wsExcel.GetRow(vLinesNo).GetCell(vColNo).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(vColNo).SetCellValue(vTempStr);
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, vColNo, vColNo));
                                wsExcel.GetRow(vLinesNo).GetCell(vColNo).CellStyle = csData;
                            }
                        }
                    }
                    vLinesNo += 2;
                }
                if (dtExcel.Rows.Count > 0)
                {
                    vLinesNo++;
                    wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                    if (vLinesNo < 13)
                    {
                        for (int ik = vLinesNo + 1; ik < 13; ik++)
                        {
                            wsExcel.CreateRow(ik);
                        }
                        vLinesNo = 13;
                    }
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("上開送審人事是否有當");
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 0, 4));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("副");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("業");
                    wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(11).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(12).SetCellValue("主");
                    wsExcel.GetRow(vLinesNo).GetCell(12).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(11).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(12).SetCellValue("辦");
                    wsExcel.GetRow(vLinesNo).GetCell(12).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("核示");
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, 4));
                    wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellValue("課");
                    wsExcel.GetRow(vLinesNo).GetCell(11).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellValue("長");
                    wsExcel.GetRow(vLinesNo).GetCell(11).CellStyle = csData_R;
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
                        Response.Write("alert('" + eMessage.Message + "')");
                        Response.Write("</" + "Script>");
                        //throw;
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('無符合條件資料')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        private void ExportExcel_A004(string fWhereStr, string fReportTitle, string fFileName)
        {
            string vSelectStr = "select a.AlterWord, a.DecDate, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where ClassNo = a.AlterType and fkey = '人事異動檔      ALTERS          altertype') as AlterType_C, " + Environment.NewLine +
                                "       a.EmpNo, a.[Name], (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        type' and ClassNo = a.Type) as Type_C, " + Environment.NewLine +
                                "       (select [Name] from Department where DepNo = a.DepNo) as DepName, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        TITLE' and ClassNo = a.Title) as Title_C, " + Environment.NewLine +
                                "       a.ALRemark7, a.InureDate, e.AssumeDay, e.BirthDay " + Environment.NewLine +
                                "  from Alters as a left join Employee as e on e.EmpNo = a.EmpNo " + Environment.NewLine +
                                fWhereStr +
                                " order by a.AlterWord, a.AlterNo ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DateTime vTempDate;
            string vTempStr = "";
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

            HSSFCellStyle csTitle_R = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle_R.VerticalAlignment = VerticalAlignment.Center;
            csTitle_R.Alignment = HorizontalAlignment.Right; //水平靠右
            csTitle_R.SetFont(fontTitle);

            //設定資料內容欄位的格式
            HSSFCellStyle csData = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            HSSFFont fontData = (HSSFFont)wbExcel.CreateFont();
            fontData.FontHeightInPoints = 12;
            csData.SetFont(fontData);

            HSSFCellStyle csData_R = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_R.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csData_R.Alignment = HorizontalAlignment.Right; //水平靠右
            csData_R.SetFont(fontData);

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
            string vAlterWord = "";
            string vDecDate = "";
            string vPageName = "";
            int vLinesNo = 0;
            int vColNo = 0;
            int vPageNo = 1;

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daExcel = new SqlDataAdapter(vSelectStr, connExcel);
                connExcel.Open();
                DataTable dtExcel = new DataTable();
                daExcel.Fill(dtExcel);
                for (int i = 0; i < dtExcel.Rows.Count; i++)
                {
                    if ((dtExcel.Rows[i]["AlterWord"].ToString().Trim() != vAlterWord) || (vLinesNo == 8)) //人令字號不同或是已經寫入五筆資料時時
                    {
                        if (vLinesNo > 0) //列數大於 0 表示前面還有一頁，所以在這一頁要放頁尾
                        {
                            wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                            if (vLinesNo < 8)
                            {
                                for (int ik = vLinesNo; ik < 8; ik++)
                                {
                                    wsExcel.CreateRow(ik);
                                }
                                vLinesNo = 8;
                            }
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("上開送審人事是否有當");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 0, 3));
                            wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("副");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("業");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("主");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("辦");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("核示");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, 3));
                            wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("課");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("長");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                        }
                        vDecDate = PF.TransDateString(dtExcel.Rows[i]["DecDate"].ToString().Trim(), "C2");
                        vLinesNo = 0; //列數先歸零
                        if (dtExcel.Rows[i]["AlterWord"].ToString().Trim() == vAlterWord)
                        {
                            vPageNo++;
                            vPageName = vAlterWord + "_" + vPageNo.ToString();
                        }
                        else
                        {
                            vAlterWord = dtExcel.Rows[i]["AlterWord"].ToString().Trim();
                            vPageName = vAlterWord;
                            vPageNo = 1;
                        }
                        //wsExcel = (HSSFSheet)wbExcel.CreateSheet(vAlterWord); //根據不同的人令字號產生不同的工作表
                        wsExcel = (HSSFSheet)wbExcel.CreateSheet(vPageName); //根據不同的人令字號產生不同的工作表
                        //產生標頭欄位
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(fReportTitle);
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 0, 6));
                        wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("桃汽客人字第 " + vAlterWord + " 號");
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 7, 9));
                        wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue(vDecDate);
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 7, 9));
                        wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int ci = 0; ci < dtExcel.Columns.Count; ci++)
                        {
                            switch (dtExcel.Columns[ci].ColumnName.Trim())
                            {
                                default:
                                    vHeaderText = "";
                                    vColNo = -1;
                                    break;
                                case "AlterType_C":
                                    vHeaderText = "區分";
                                    vColNo = 0;
                                    break;
                                case "EmpNo":
                                    vHeaderText = "員工編號";
                                    vColNo = 1;
                                    break;
                                case "Name":
                                    vHeaderText = "姓名";
                                    vColNo = 2;
                                    break;
                                case "Type_C":
                                    vHeaderText = "人員類別";
                                    vColNo = 3;
                                    break;
                                case "DepName":
                                    vHeaderText = "部門";
                                    vColNo = 4;
                                    break;
                                case "Title_C":
                                    vHeaderText = "職稱";
                                    vColNo = 5;
                                    break;
                                case "ALRemark7":
                                    vHeaderText = "原因";
                                    vColNo = 6;
                                    break;
                                case "InureDate":
                                    vHeaderText = "生效日期";
                                    vColNo = 7;
                                    break;
                                case "AssumeDay":
                                    vHeaderText = "到職日期";
                                    vColNo = 8;
                                    break;
                                case "BirthDay":
                                    vHeaderText = "出生年月";
                                    vColNo = 9;
                                    break;
                            }
                            if (vColNo != -1)
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(vColNo).SetCellValue(vHeaderText);
                                wsExcel.GetRow(vLinesNo).GetCell(vColNo).CellStyle = csData;
                            }
                        }
                        vLinesNo++;
                    }
                    //開始寫入資料
                    wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                    wsExcel.CreateRow(vLinesNo);
                    for (int ci = 0; ci < dtExcel.Columns.Count; ci++)
                    {
                        ///*
                        switch (dtExcel.Columns[ci].ColumnName.Trim())
                        {
                            default:
                                vColNo = -1;
                                break;
                            case "AlterType_C":
                                vTempStr = dtExcel.Rows[i]["AlterType_C"].ToString().Trim();
                                vColNo = 0;
                                break;
                            case "EmpNo":
                                vTempStr = dtExcel.Rows[i]["EmpNo"].ToString().Trim();
                                vColNo = 1;
                                break;
                            case "Name":
                                vTempStr = dtExcel.Rows[i]["Name"].ToString().Trim();
                                vColNo = 2;
                                break;
                            case "Type_C":
                                vTempStr = dtExcel.Rows[i]["Type_C"].ToString().Trim();
                                vColNo = 3;
                                break;
                            case "DepName":
                                vTempStr = dtExcel.Rows[i]["DepName"].ToString().Trim();
                                vColNo = 4;
                                break;
                            case "Title_C":
                                vTempStr = dtExcel.Rows[i]["Title_C"].ToString().Trim();
                                vColNo = 5;
                                break;
                            case "ALRemark7":
                                vTempStr = dtExcel.Rows[i]["ALRemark7"].ToString().Trim();
                                vColNo = 6;
                                break;
                            case "InureDate":
                                vTempDate = DateTime.Parse(dtExcel.Rows[i]["InureDate"].ToString().Trim());
                                vTempStr = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                                vColNo = 7;
                                break;
                            case "AssumeDay":
                                vTempDate = DateTime.Parse(dtExcel.Rows[i]["AssumeDay"].ToString().Trim());
                                vTempStr = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                                vColNo = 8;
                                break;
                            case "BirthDay":
                                vTempDate = DateTime.Parse(dtExcel.Rows[i]["BirthDay"].ToString().Trim());
                                vTempStr = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                                vColNo = 9;
                                break;
                        }
                        if (vColNo != -1)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(vColNo);
                            wsExcel.GetRow(vLinesNo).GetCell(vColNo).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(vColNo).SetCellValue(vTempStr);
                            wsExcel.GetRow(vLinesNo).GetCell(vColNo).CellStyle = csData;
                        }
                    }
                    vLinesNo++;
                }
                if (dtExcel.Rows.Count > 0)
                {
                    vLinesNo++;
                    wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                    if (vLinesNo < 8)
                    {
                        for (int ik = vLinesNo + 1; ik < 8; ik++)
                        {
                            wsExcel.CreateRow(ik);
                        }
                        vLinesNo = 8;
                    }
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("上開送審人事是否有當");
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 0, 3));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("副");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("業");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("主");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("辦");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("核示");
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, 3));
                    wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("課");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("長");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
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
                        Response.Write("alert('" + eMessage.Message + "')");
                        Response.Write("</" + "Script>");
                        //throw;
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('無符合條件資料')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        private void ExportExcel_A005(string fWhereStr, string fReportTitle, string fFileName)
        {
            string vSelectStr = "select a.AlterWord, a.DecDate, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where ClassNo = a.AlterType and fkey = '人事異動檔      ALTERS          altertype') as AlterType_C, " + Environment.NewLine +
                                "       a.EmpNo, a.[Name], (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        type' and ClassNo = a.Type) as Type_C, " + Environment.NewLine +
                                "       (select [Name] from Department where DepNo = a.DepNo) as DepName, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        TITLE' and ClassNo = a.Title) as Title_C, " + Environment.NewLine +
                                "       a.ALRemark7, a.InureDate, e.AssumeDay, e.BirthDay " + Environment.NewLine +
                                "  from Alters as a left join Employee as e on e.EmpNo = a.EmpNo " + Environment.NewLine +
                                fWhereStr +
                                " order by a.AlterWord, a.AlterNo ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DateTime vTempDate;
            string vTempStr = "";
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

            HSSFCellStyle csTitle_R = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle_R.VerticalAlignment = VerticalAlignment.Center;
            csTitle_R.Alignment = HorizontalAlignment.Right; //水平靠右
            csTitle_R.SetFont(fontTitle);

            //設定資料內容欄位的格式
            HSSFCellStyle csData = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            HSSFFont fontData = (HSSFFont)wbExcel.CreateFont();
            fontData.FontHeightInPoints = 12;
            csData.SetFont(fontData);

            HSSFCellStyle csData_R = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_R.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csData_R.Alignment = HorizontalAlignment.Right; //水平靠右
            csData_R.SetFont(fontData);

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
            string vAlterWord = "";
            string vDecDate = "";
            string vPageName = "";
            int vLinesNo = 0;
            int vColNo = 0;
            int vPageNo = 1;

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daExcel = new SqlDataAdapter(vSelectStr, connExcel);
                connExcel.Open();
                DataTable dtExcel = new DataTable();
                daExcel.Fill(dtExcel);
                for (int i = 0; i < dtExcel.Rows.Count; i++)
                {
                    if ((dtExcel.Rows[i]["AlterWord"].ToString().Trim() != vAlterWord) || (vLinesNo == 8)) //人令字號不同或是已經寫入五筆資料時時
                    {
                        if (vLinesNo > 0) //列數大於 0 表示前面還有一頁，所以在這一頁要放頁尾
                        {
                            wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                            if (vLinesNo < 8)
                            {
                                for (int ik = vLinesNo; ik < 8; ik++)
                                {
                                    wsExcel.CreateRow(ik);
                                }
                                vLinesNo = 8;
                            }
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("上開送審人事是否有當");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 0, 3));
                            wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("副");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("業");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("主");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("總");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("務");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("辦");
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("核示");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, 3));
                            wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("經");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("課");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("理");
                            wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                            wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("長");
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                        }
                        vDecDate = PF.TransDateString(dtExcel.Rows[i]["DecDate"].ToString().Trim(), "C2");
                        vLinesNo = 0; //列數先歸零
                        if (dtExcel.Rows[i]["AlterWord"].ToString().Trim() == vAlterWord)
                        {
                            vPageNo++;
                            vPageName = vAlterWord + "_" + vPageNo.ToString();
                        }
                        else
                        {
                            vAlterWord = dtExcel.Rows[i]["AlterWord"].ToString().Trim();
                            vPageName = vAlterWord;
                            vPageNo = 1;
                        }
                        //wsExcel = (HSSFSheet)wbExcel.CreateSheet(vAlterWord); //根據不同的人令字號產生不同的工作表
                        wsExcel = (HSSFSheet)wbExcel.CreateSheet(vPageName); //根據不同的人令字號產生不同的工作表
                        //產生標頭欄位
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(fReportTitle);
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 0, 6));
                        wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("桃汽客人字第 " + vAlterWord + " 號");
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 7, 9));
                        wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue(vDecDate);
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 7, 9));
                        wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int ci = 0; ci < dtExcel.Columns.Count; ci++)
                        {
                            switch (dtExcel.Columns[ci].ColumnName.Trim())
                            {
                                default:
                                    vHeaderText = "";
                                    vColNo = -1;
                                    break;
                                case "AlterType_C":
                                    vHeaderText = "區分";
                                    vColNo = 0;
                                    break;
                                case "EmpNo":
                                    vHeaderText = "員工編號";
                                    vColNo = 1;
                                    break;
                                case "Name":
                                    vHeaderText = "姓名";
                                    vColNo = 2;
                                    break;
                                case "Type_C":
                                    vHeaderText = "人員類別";
                                    vColNo = 3;
                                    break;
                                case "DepName":
                                    vHeaderText = "部門";
                                    vColNo = 4;
                                    break;
                                case "Title_C":
                                    vHeaderText = "職稱";
                                    vColNo = 5;
                                    break;
                                case "ALRemark7":
                                    vHeaderText = "原因";
                                    vColNo = 6;
                                    break;
                                case "ALRemark8":
                                    vHeaderText = "審核意見";
                                    vColNo = 7;
                                    break;
                                case "InureDate":
                                    vHeaderText = "生效日期";
                                    vColNo = 8;
                                    break;
                                case "AssumeDay":
                                    vHeaderText = "到職日期";
                                    vColNo = 9;
                                    break;
                            }
                            if (vColNo != -1)
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(vColNo).SetCellValue(vHeaderText);
                                wsExcel.GetRow(vLinesNo).GetCell(vColNo).CellStyle = csData;
                            }
                        }
                        vLinesNo++;
                    }
                    //開始寫入資料
                    wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                    wsExcel.CreateRow(vLinesNo);
                    for (int ci = 0; ci < dtExcel.Columns.Count; ci++)
                    {
                        ///*
                        switch (dtExcel.Columns[ci].ColumnName.Trim())
                        {
                            default:
                                vColNo = -1;
                                break;
                            case "AlterType_C":
                                vTempStr = dtExcel.Rows[i]["AlterType_C"].ToString().Trim();
                                vColNo = 0;
                                break;
                            case "EmpNo":
                                vTempStr = dtExcel.Rows[i]["EmpNo"].ToString().Trim();
                                vColNo = 1;
                                break;
                            case "Name":
                                vTempStr = dtExcel.Rows[i]["Name"].ToString().Trim();
                                vColNo = 2;
                                break;
                            case "Type_C":
                                vTempStr = dtExcel.Rows[i]["Type_C"].ToString().Trim();
                                vColNo = 3;
                                break;
                            case "DepName":
                                vTempStr = dtExcel.Rows[i]["DepName"].ToString().Trim();
                                vColNo = 4;
                                break;
                            case "Title_C":
                                vTempStr = dtExcel.Rows[i]["Title_C"].ToString().Trim();
                                vColNo = 5;
                                break;
                            case "ALRemark7":
                                vTempStr = dtExcel.Rows[i]["ALRemark7"].ToString().Trim();
                                vColNo = 6;
                                break;
                            case "ALRemark8":
                                vTempStr = dtExcel.Rows[i]["ALRemark8"].ToString().Trim();
                                vColNo = 7;
                                break;
                            case "InureDate":
                                vTempDate = DateTime.Parse(dtExcel.Rows[i]["InureDate"].ToString().Trim());
                                vTempStr = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                                vColNo = 8;
                                break;
                            case "AssumeDay":
                                vTempDate = DateTime.Parse(dtExcel.Rows[i]["AssumeDay"].ToString().Trim());
                                vTempStr = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                                vColNo = 9;
                                break;
                        }
                        if (vColNo != -1)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(vColNo);
                            wsExcel.GetRow(vLinesNo).GetCell(vColNo).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(vColNo).SetCellValue(vTempStr);
                            wsExcel.GetRow(vLinesNo).GetCell(vColNo).CellStyle = csData;
                        }
                    }
                    vLinesNo++;
                }
                if (dtExcel.Rows.Count > 0)
                {
                    vLinesNo++;
                    wsExcel = (HSSFSheet)wbExcel.GetSheet(vPageName);
                    if (vLinesNo < 8)
                    {
                        for (int ik = vLinesNo + 1; ik < 8; ik++)
                        {
                            wsExcel.CreateRow(ik);
                        }
                        vLinesNo = 8;
                    }
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("上開送審人事是否有當");
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 0, 3));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("副");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("業");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("主");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("總");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("務");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("辦");
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("核示");
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, 3));
                    wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("經");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("課");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("理");
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData_R;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("長");
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData_R;
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
                        Response.Write("alert('" + eMessage.Message + "')");
                        Response.Write("</" + "Script>");
                        //throw;
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('無符合條件資料')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        protected void bbCancel_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}