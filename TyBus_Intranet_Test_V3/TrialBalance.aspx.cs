using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class TrialBalance : System.Web.UI.Page
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
        private DateTime vToday = DateTime.Today;
        private string vCompanyName = "桃園汽車客運股份有限公司";

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
                    string vTempDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCalDate_Start_Search.ClientID;
                    string vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCalDate_Start_Search.Attributes["onClick"] = vTempDateScript;

                    vTempDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCalDate_End_Search.ClientID;
                    vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCalDate_End_Search.Attributes["onClick"] = vTempDateScript;

                    if (!IsPostBack)
                    {
                        eCalDate_Start_Search.Text = PF.GetMonthFirstDay(DateTime.Today, "C");
                        eCalDate_End_Search.Text = PF.GetMonthLastDay(DateTime.Today, "C");
                        plReport.Visible = false;
                        //plReport2.Visible = false;
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
            string vResultStr = "";
            //起始日期
            DateTime vDate_Start_D = (eCalDate_Start_Search.Text.Trim() != "") ? DateTime.Parse(eCalDate_Start_Search.Text.Trim()) : DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today, "C"));
            string vDate_Start = (eCalDate_Start_Search.Text.Trim() != "") ? PF.TransDateString(eCalDate_Start_Search.Text.Trim(), "B") : PF.GetMonthFirstDay(DateTime.Today, "B");
            //結束日期
            DateTime vDate_End_D = (eCalDate_End_Search.Text.Trim() != "") ? DateTime.Parse(eCalDate_End_Search.Text.Trim()) : DateTime.Parse(PF.GetMonthLastDay(DateTime.Today, "C"));
            string vDate_End = (eCalDate_End_Search.Text.Trim() != "") ? PF.TransDateString(eCalDate_End_Search.Text.Trim(), "B") : PF.GetMonthLastDay(DateTime.Today, "B");
            //前一年年底
            string vDate_LastYearLastDay = (DateTime.Parse(eCalDate_Start_Search.Text.Trim()).Year - 1).ToString("D4") + "/12/31";
            //下月首日
            string vDate_NextMonthFirstDay = (DateTime.Parse(eCalDate_End_Search.Text.Trim()).AddDays(1)).Year.ToString("D4") + "/" + (DateTime.Parse(eCalDate_End_Search.Text.Trim()).AddDays(1)).ToString("MM/dd");
            //存貨科目
            string vSQLStr = "select Desscription from AC_CLS where PK='fin2000' and (class='存    貨' or class='商品存貨' or class='存貨')";
            string vSubject_Stock = PF.GetValue(vConnStr, vSQLStr, "Desscription");

            string vWStr_Subject = ((eSubjectNo_Start_Search.Text.Trim() != "") && (eSubjectNo_End_Search.Text.Trim() != "")) ? " and a.[Subject] between '" + eSubjectNo_Start_Search.Text.Trim() + "' and '" + eSubjectNo_End_Search.Text.Trim() + "' " :
                                   ((eSubjectNo_Start_Search.Text.Trim() != "") && (eSubjectNo_End_Search.Text.Trim() == "")) ? " and a.[Subject] like '" + eSubjectNo_Start_Search.Text.Trim() + "%' " :
                                   ((eSubjectNo_Start_Search.Text.Trim() == "") && (eSubjectNo_End_Search.Text.Trim() != "")) ? " and a.[Subject] like '" + eSubjectNo_End_Search.Text.Trim() + "%' " : "";
            vResultStr = "select * from (" + Environment.NewLine +
                         "select t.[Subject] 科目編號, t.[Name] 科目名稱, " + Environment.NewLine +
                         "       case when sum(t.PBalance) > 0 then sum(t.PBalance) else 0 end 上月借方結存, " + Environment.NewLine +
                         "       case when sum(t.PBalance) < 0 then ABS(sum(t.PBalance)) else 0 end 上月貸方結存, " + Environment.NewLine +
                         "       sum(t.SDebit) 本月借方合計, sum(t.SCredit) 本月貸方合計, " + Environment.NewLine +
                         "       case when sum(t.PBalance) + sum(t.SDebit) - sum(t.SCredit) > 0 then sum(t.PBalance) + sum(t.SDebit) - sum(t.SCredit) else 0 end 本月借方結存, " + Environment.NewLine +
                         "       case when sum(t.PDebit) - sum(t.PCredit) + sum(t.SDebit) - sum(t.SCredit) < 0 then ABS(sum(t.PDebit) - sum(t.PCredit) + sum(t.SDebit) - sum(t.SCredit)) else 0 end 本月貸方結存 " + Environment.NewLine +
                         "  from (" + Environment.NewLine +
                         "        select a.[Subject] as subject, b.[Name] as name, sum(a.debit) as sdebit, sum(a.credit) as scredit, " + Environment.NewLine +
                         "               sum(a.debit) - sum(a.credit) as balance, sum(a.debit) - sum(a.debit) as pdebit, " + Environment.NewLine +
                         "               sum(a.debit) - sum(a.debit) as pcredit, sum(a.debit) - sum(a.debit) as PBalance " + Environment.NewLine +
                         "          from Account a, AC_Subject b " + Environment.NewLine +
                         "         where a.[Subject] = B.[subject] " + Environment.NewLine +
                         "           and a.REC_Date >= '" + vDate_Start + "' and a.REC_Date < '" + vDate_NextMonthFirstDay + "' " + Environment.NewLine +
                         "           and Type <> '6' and a.VoidMan IS NULL " + vWStr_Subject + Environment.NewLine +
                         "         group by a.[Subject], b.[Name] " + Environment.NewLine +
                         "         union all " + Environment.NewLine +
                         "        select z.[Subject], z.[Name], cast(0 as float) sdebit, cast(0 as float) scredit, cast(0 as float) balance, " + Environment.NewLine +
                         "               z.pdebit, z.pcredit, z.PBalance " + Environment.NewLine +
                         "          from (" + Environment.NewLine +
                         "                Select a.[Subject] as subject, b.[Name] as name, " + Environment.NewLine +
                         "                       sum(a.Debit) as PDebit, sum(a.Credit) as PCredit, " + Environment.NewLine +
                         "                       sum(a.Debit) - sum(a.Credit) as PBalance " + Environment.NewLine +
                         "                  from TOTAL_ACCOUNT a, AC_Subject b " + Environment.NewLine +
                         "                 where a.[Subject] = B.[Subject] " + Environment.NewLine +
                         "                   and a.[Subject] <> '3170' " + Environment.NewLine +
                         "                   and a.[Subject] <> '1151' " + Environment.NewLine +
                         "                   and a.yymm_Date < '" + vDate_Start + "' " + Environment.NewLine +
                         "                   and a.[Subject] < '4%' " + vWStr_Subject + Environment.NewLine +
                         "                 group by a.[Subject], b.[Name] " + Environment.NewLine +
                         "                 union all " + Environment.NewLine +
                         "                select a.[Subject] as [Subject], b.[Name] as [Name], " + Environment.NewLine +
                         "                       sum(a.Debit) as PDebit, sum(a.Credit) as PCredit, sum(a.Debit) - sum(a.Credit) as PBalance " + Environment.NewLine +
                         "                  from TOTAL_ACCOUNT a, AC_Subject b " + Environment.NewLine +
                         "                 where a.[Subject] = b.[Subject] " + Environment.NewLine +
                         "                   and a.[Subject] <> '3170' " + Environment.NewLine +
                         "                   and a.yymm_Date > '" + vDate_LastYearLastDay + "' " + Environment.NewLine +
                         "                   and a.yymm_Date < '" + vDate_Start + "' " + Environment.NewLine +
                         "                   and a.[Subject] > '4%' " + vWStr_Subject + Environment.NewLine +
                         "                 group by a.[Subject], b.[Name] " + Environment.NewLine +
                         "                 union all " + Environment.NewLine +
                         "                select '" + vSubject_Stock + "' as [Subject], b.[Name] as [Name], sum(a.Debit) as PDebit, sum(a.Credit) as PCredit, " + Environment.NewLine +
                         "                       sum(a.Debit) - sum(a.Credit) as PBalance " + Environment.NewLine +
                         "                  from TOTAL_ACCOUNT a, AC_Subject b " + Environment.NewLine +
                         "                 where a.[Subject] = b.[Subject] " + Environment.NewLine +
                         "                   and a.[Subject] = '" + vSubject_Stock + "' " + Environment.NewLine +
                         "                   and Credit = 0 and Debit <> 0 " + Environment.NewLine +
                         "                 group by a.[Subject], b.[Name] " + Environment.NewLine +
                         "               ) z " + Environment.NewLine +
                         "       ) t " + Environment.NewLine +
                         " group by t.[Subject], t.[Name] " + Environment.NewLine +
                         ") y " + Environment.NewLine +
                         " where y.上月借方結存 > 0 or y.上月貸方結存 > 0 or y.本月借方合計 > 0 or y.本月貸方合計 > 0 " + Environment.NewLine +
                         " order by y.科目編號";
            return vResultStr;
        }

        protected void bbTransToExcel_Click(object sender, EventArgs e)
        {
            Server.ScriptTimeout = 1440;

            string vSQLStr = "";
            //起始日期
            DateTime vDate_Start_D = (eCalDate_Start_Search.Text.Trim() != "") ? DateTime.Parse(eCalDate_Start_Search.Text.Trim()) : DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today, "C"));
            string vDate_Start = (eCalDate_Start_Search.Text.Trim() != "") ? PF.TransDateString(eCalDate_Start_Search.Text.Trim(), "B") : PF.GetMonthFirstDay(DateTime.Today, "B");
            //結束日期
            DateTime vDate_End_D = (eCalDate_End_Search.Text.Trim() != "") ? DateTime.Parse(eCalDate_End_Search.Text.Trim()) : DateTime.Parse(PF.GetMonthLastDay(DateTime.Today, "C"));
            string vDate_End = (eCalDate_End_Search.Text.Trim() != "") ? PF.TransDateString(eCalDate_End_Search.Text.Trim(), "B") : PF.GetMonthLastDay(DateTime.Today, "B");
            //前一年年底
            string vDate_LastYearLastDay = (DateTime.Parse(eCalDate_Start_Search.Text.Trim()).Year - 1).ToString("D4") + "/12/31";
            //下月首日
            string vDate_NextMonthFirstDay = (DateTime.Parse(eCalDate_End_Search.Text.Trim()).AddDays(1)).Year.ToString("D4") + "/" + (DateTime.Parse(eCalDate_End_Search.Text.Trim()).AddDays(1)).ToString("MM/dd");
            //當年度首日
            string vDate_ThisYearFirstDay = PF.GetYearFirstDay(DateTime.Parse(eCalDate_Start_Search.Text.Trim()), "B");
            //當月首日
            string vDate_ThisMonthFirst = PF.GetMonthFirstDay(DateTime.Parse(eCalDate_Start_Search.Text.Trim()), "B");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelectStr();
            using (SqlConnection connTrialBalance = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daTrialBalance = new SqlDataAdapter(vSelectStr, connTrialBalance);
                connTrialBalance.Open();
                DataTable dtTrialBalance = new DataTable();
                daTrialBalance.Fill(dtTrialBalance);
                if (dtTrialBalance.Rows.Count > 0)
                {
                    //Excel 檔案
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
                    fontTitle.FontHeightInPoints = 20; //字體大小
                    csTitle.SetFont(fontTitle);

                    //子標題格式
                    HSSFCellStyle csSubTitle = (HSSFCellStyle)wbExcel.CreateCellStyle();
                    csSubTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                    csSubTitle.Alignment = HorizontalAlignment.Center; //水平置中

                    //要靠右的欄位套用格式
                    HSSFCellStyle csRightAligment = (HSSFCellStyle)wbExcel.CreateCellStyle();
                    csRightAligment.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                    csRightAligment.Alignment = HorizontalAlignment.Right; //水平靠右

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

                    //換頁用的欄位值
                    string vGroupText = "";
                    int vLinesNo = 0;
                    //先把 "累進試算表" 寫到第一個工作表
                    vGroupText = "累進試算表";
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vGroupText);
                    //建立標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < dtTrialBalance.Columns.Count; i++)
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(dtTrialBalance.Columns[i].Caption);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    //開始寫入資料
                    for (int j = 0; j < dtTrialBalance.Rows.Count; j++)
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int k = 0; k < dtTrialBalance.Columns.Count; k++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(k);
                            if (k < 2)
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(k).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(k).SetCellValue(dtTrialBalance.Rows[j][k].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(k).CellStyle = csData;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(k).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(k).SetCellValue(double.Parse(dtTrialBalance.Rows[j][k].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(k).CellStyle = csData;
                            }
                        }
                    }
                    //寫入小計欄
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(1);
                    wsExcel.GetRow(vLinesNo).GetCell(1).SetCellValue("合計：");
                    wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csRightAligment;
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellFormula("sum(C2:C" + vLinesNo.ToString() + ")");
                    wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellFormula("sum(D2:D" + vLinesNo.ToString() + ")");
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellFormula("sum(E2:E" + vLinesNo.ToString() + ")");
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellFormula("sum(F2:F" + vLinesNo.ToString() + ")");
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellFormula("sum(G2:G" + vLinesNo.ToString() + ")");
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellFormula("sum(H2:H" + vLinesNo.ToString() + ")");
                    //再來依據不同的科目取回分類帳
                    string vSubjectNo = "";
                    string vSubjectName = "";
                    string vSelectStr_Sub = "";
                    string vPDebit = "0";
                    string vPCredit = "0";
                    string vPBalanceStr = "";
                    double vDebit_Detail = 0;
                    double vCredit_Detail = 0;
                    double vPBalance = 0;
                    string vRECDate = "";

                    for (int vSubjectIndex = 0; vSubjectIndex < dtTrialBalance.Rows.Count; vSubjectIndex++)
                    {
                        vSubjectNo = dtTrialBalance.Rows[vSubjectIndex][0].ToString().Trim();
                        vSubjectName = dtTrialBalance.Rows[vSubjectIndex][1].ToString().Trim();
                        vLinesNo = 0;
                        vSQLStr = "select sum(SDebit) SDebit " + Environment.NewLine +
                                  "  from ( " + Environment.NewLine +
                                  "        select sum(a.Debit) as SDebit, sum(a.Credit) as SCredit " + Environment.NewLine +
                                  "          from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                                  "         where a.[Subject] = b.[Subject] and a.[Subject] = '" + vSubjectNo + "' " + Environment.NewLine +
                                  "           and a.YYMM_DATE < '" + vDate_Start + "' and a.[Subject] < '4%' " + Environment.NewLine +
                                  "         group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                                  "         union all " + Environment.NewLine +
                                  "        select sum(a.Debit) as SDebit, sum(a.Credit) as SCredit " + Environment.NewLine +
                                  "          from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                                  "         where a.[Subject] = b.[Subject] and a.[Subject] = '" + vSubjectNo + "' " + Environment.NewLine +
                                  "           and a.YYMM_DATE < '" + vDate_Start + "' and a.YYMM_DATE >= '" + vDate_ThisYearFirstDay + "' " + Environment.NewLine +
                                  "           and a.[Subject] > '4%' " + Environment.NewLine +
                                  "         group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                                  "         union all " + Environment.NewLine +
                                  "        select sum(a.Debit) as SDebit, sum(a.Credit) as SCredit " + Environment.NewLine +
                                  "          from Account a , AC_Subject b " + Environment.NewLine +
                                  "         where a.[Subject] = b.[Subject] and a.[Subject] = '" + vSubjectNo + "' " + Environment.NewLine +
                                  "           and a.VoidDate IS NULL " + Environment.NewLine +
                                  "           and a.REC_Date >= '" + vDate_ThisMonthFirst + "' and a.REC_Date < '" + vDate_Start + "' " + Environment.NewLine +
                                  "         group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                                  "        ) t";
                        vPDebit = PF.GetValue(vConnStr, vSQLStr, "SDebit");
                        if (vPDebit == "")
                        {
                            vPDebit = "0";
                        }
                        vSQLStr = "select sum(SCredit) SCredit " + Environment.NewLine +
                                  "  from ( " + Environment.NewLine +
                                  "        select sum(a.Debit) as SDebit, sum(a.Credit) as SCredit " + Environment.NewLine +
                                  "          from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                                  "         where a.[Subject] = b.[Subject] and a.[Subject] = '" + vSubjectNo + "' " + Environment.NewLine +
                                  "           and a.YYMM_DATE < '" + vDate_Start + "' and a.[Subject] < '4%' " + Environment.NewLine +
                                  "         group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                                  "         union all " + Environment.NewLine +
                                  "        select sum(a.Debit) as SDebit, sum(a.Credit) as SCredit " + Environment.NewLine +
                                  "          from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                                  "         where a.[Subject] = b.[Subject] and a.[Subject] = '" + vSubjectNo + "' " + Environment.NewLine +
                                  "           and a.YYMM_DATE < '" + vDate_Start + "' and a.YYMM_DATE >= '" + vDate_ThisYearFirstDay + "' " + Environment.NewLine +
                                  "           and a.[Subject] > '4%' " + Environment.NewLine +
                                  "         group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                                  "         union all " + Environment.NewLine +
                                  "        select sum(a.Debit) as SDebit, sum(a.Credit) as SCredit " + Environment.NewLine +
                                  "          from Account a , AC_Subject b " + Environment.NewLine +
                                  "         where a.[Subject] = b.[Subject] and a.[Subject] = '" + vSubjectNo + "' " + Environment.NewLine +
                                  "           and a.VoidDate IS NULL " + Environment.NewLine +
                                  "           and a.REC_Date >= '" + vDate_ThisMonthFirst + "' and a.REC_Date < '" + vDate_Start + "' " + Environment.NewLine +
                                  "         group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                                  "        ) t";
                        vPCredit = PF.GetValue(vConnStr, vSQLStr, "SCredit");
                        if (vPCredit == "")
                        {
                            vPCredit = "0";
                        }
                        vSelectStr_Sub = "select a.[No], a.Product, a.REC_Date, a.Item, a.[Subject], a.[Target], a.Meno, " + Environment.NewLine +
                                         "       a.Debit, a.Credit, b.[Name], cast(" + vPDebit + " as float) as PDebit, cast(" + vPCredit + " as float) as PCredit, " + Environment.NewLine +
                                         "       For_Type, For_Debit, For_Credit, Department, (select [Name] from Department where DepNo = a.Department) DepName, " + Environment.NewLine +
                                         "       b.IsDebit as DebitType, a.Memo2 " + Environment.NewLine +
                                         "  from Account a, AC_Subject b " + Environment.NewLine +
                                         " where a.[Type] <> '6' and a.[Subject] = b.[Subject] " + Environment.NewLine +
                                         "   and a.REC_Date >= '" + vDate_Start + "' and a.REC_Date < '" + vDate_NextMonthFirstDay + "' " + Environment.NewLine +
                                         "   and b.[Subject] ='" + vSubjectNo + "' " + Environment.NewLine +
                                         "   and a.VoidDate IS NULL " + Environment.NewLine +
                                         " order by a.REC_Date";
                        vPBalance = ((vSubjectNo.Trim().Substring(0, 1) == "1") ||
                                             (vSubjectNo.Trim().Substring(0, 1) == "5") ||
                                             (vSubjectNo.Trim().Substring(0, 1) == "8") ||
                                             (vSubjectNo.Trim().Substring(0, 1) == "9") ||
                                             (vSubjectNo.Trim().Substring(0, 2) == "62")) ?
                                             double.Parse(vPDebit) - double.Parse(vPCredit) :
                                             double.Parse(vPCredit) - double.Parse(vPDebit);
                        vPBalanceStr = (vPBalance == 0) ? "" :
                                       (((vSubjectNo.Trim().Substring(0, 1) == "1") ||
                                         (vSubjectNo.Trim().Substring(0, 1) == "5") ||
                                         (vSubjectNo.Trim().Substring(0, 1) == "8") ||
                                         (vSubjectNo.Trim().Substring(0, 1) == "9") ||
                                         (vSubjectNo.Trim().Substring(0, 2) == "62")) &&
                                        (vPBalance > 0)) ? "[ 借 ]" :
                                       (((vSubjectNo.Trim().Substring(0, 1) == "2") ||
                                         (vSubjectNo.Trim().Substring(0, 1) == "3") ||
                                         (vSubjectNo.Trim().Substring(0, 1) == "4") ||
                                         (vSubjectNo.Trim().Substring(0, 1) == "7")) &&
                                        (vPBalance < 0)) ? "[ 借 ]" : "[ 貸 ]";
                        using (SqlConnection connLedger = new SqlConnection(vConnStr))
                        {
                            SqlDataAdapter daLedger = new SqlDataAdapter(vSelectStr_Sub, connLedger);
                            DataTable dtLedger = new DataTable();
                            daLedger.Fill(dtLedger);
                            if (dtLedger.Rows.Count > 0)
                            {
                                //建立工作表
                                wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSubjectNo + "-" + vSubjectName);
                                //建立表頭區
                                vLinesNo = 0;
                                wsExcel.CreateRow(vLinesNo);
                                wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("科目分類帳明細表");
                                //wsExcel.AutoSizeColumn(0);
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 7));
                                wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                                vLinesNo++;
                                wsExcel.CreateRow(vLinesNo);
                                wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("資料日期：" +
                                                                                    (vDate_Start_D.Year - 1911).ToString("D3") + "/" + vDate_Start_D.ToString("MM/dd") +
                                                                                    "～" +
                                                                                    (vDate_End_D.Year - 1911).ToString("D3") + "/" + vDate_End_D.ToString("MM/dd"));
                                //wsExcel.AutoSizeColumn(0);
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 2));
                                wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("科目範圍：" + vSubjectNo);
                                //wsExcel.AutoSizeColumn(3);
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 3, 4));
                                vLinesNo++;
                                wsExcel.CreateRow(vLinesNo);
                                wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("科目代號：" + vSubjectNo);
                                //wsExcel.AutoSizeColumn(0);
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 1));
                                wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("科目名稱：" + vSubjectName);
                                //wsExcel.AutoSizeColumn(2);
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, 3));
                                wsExcel.GetRow(vLinesNo).CreateCell(4);
                                wsExcel.GetRow(vLinesNo).GetCell(4).SetCellValue("期初餘額：");
                                wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csRightAligment;
                                //wsExcel.AutoSizeColumn(4);
                                wsExcel.GetRow(vLinesNo).CreateCell(5);
                                wsExcel.GetRow(vLinesNo).GetCell(5).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(5).SetCellValue(vPBalance);
                                wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue(vPBalanceStr);
                                vLinesNo++;
                                wsExcel.CreateRow(vLinesNo);
                                wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("");
                                //wsExcel.AutoSizeColumn(0);
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 7));
                                //建立標題列
                                vLinesNo++;
                                wsExcel.CreateRow(vLinesNo);
                                wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("日期");
                                wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue("傳票號碼");
                                wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("部門");
                                wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("摘要");
                                wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("借方金額");
                                wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("貸方金額");
                                wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("餘");
                                wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("本期餘額");
                                wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("摘要二");
                                wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csSubTitle;
                                //寫入內容
                                for (int vDataIndex = 0; vDataIndex < dtLedger.Rows.Count; vDataIndex++)
                                {
                                    vLinesNo++;
                                    wsExcel.CreateRow(vLinesNo);
                                    vRECDate = (DateTime.Parse(dtLedger.Rows[vDataIndex]["REC_Date"].ToString().Trim()).Year - 1911).ToString("D3") + "/" + DateTime.Parse(dtLedger.Rows[vDataIndex]["REC_Date"].ToString().Trim()).ToString("MM/dd");
                                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(vRECDate);
                                    wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue(dtLedger.Rows[vDataIndex]["No"].ToString().Trim());
                                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue(dtLedger.Rows[vDataIndex]["DepName"].ToString().Trim());
                                    wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue(dtLedger.Rows[vDataIndex]["Meno"].ToString().Trim());
                                    wsExcel.GetRow(vLinesNo).CreateCell(4);
                                    wsExcel.GetRow(vLinesNo).GetCell(4).SetCellType(CellType.Numeric);
                                    vDebit_Detail = double.Parse(dtLedger.Rows[vDataIndex]["Debit"].ToString().Trim());
                                    wsExcel.GetRow(vLinesNo).GetCell(4).SetCellValue(vDebit_Detail);
                                    wsExcel.GetRow(vLinesNo).CreateCell(5);
                                    wsExcel.GetRow(vLinesNo).GetCell(5).SetCellType(CellType.Numeric);
                                    vCredit_Detail = double.Parse(dtLedger.Rows[vDataIndex]["Credit"].ToString().Trim());
                                    wsExcel.GetRow(vLinesNo).GetCell(5).SetCellValue(vCredit_Detail);
                                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellFormula("IF((IF(G3=\"[ 借 ]\"," +
                                                                                          "       sum(E6:E" + (vLinesNo + 1).ToString() + ") - sum(F6:F" + (vLinesNo + 1).ToString() + ") + F3," +
                                                                                          "       sum(E6:E" + (vLinesNo + 1).ToString() + ") - sum(F6:F" + (vLinesNo + 1).ToString() + ") - F3)) < 0," +
                                                                                          "\"貸\"," +
                                                                                          "\"借\")");
                                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csRightAligment;
                                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellFormula("ABS(IF(G3=\"[ 借 ]\"," +
                                                                                          "    sum(E6:E" + (vLinesNo + 1).ToString() + ") - sum(F6:F" + (vLinesNo + 1).ToString() + ") + F3," +
                                                                                          "    sum(E6:E" + (vLinesNo + 1).ToString() + ") - sum(F6:F" + (vLinesNo + 1).ToString() + ") - F3))");
                                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue(dtLedger.Rows[vDataIndex]["Memo2"].ToString().Trim());
                                }
                                //寫入小計
                                vLinesNo++;
                                wsExcel.CreateRow(vLinesNo);
                                wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("小計：");
                                wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellFormula("sum(E6:E" + vLinesNo.ToString() + ")");
                                wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellFormula("sum(F6:F" + vLinesNo.ToString() + ")");
                            }
                            else
                            {
                                //建立工作表
                                wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSubjectNo + "-" + vSubjectName);
                                //建立表頭區
                                vLinesNo = 0;
                                wsExcel.CreateRow(vLinesNo);
                                wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("科目分類帳明細表");
                                //wsExcel.AutoSizeColumn(0);
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 7));
                                wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                                vLinesNo++;
                                wsExcel.CreateRow(vLinesNo);
                                wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("資料日期：" +
                                                                                    (vDate_Start_D.Year - 1911).ToString("D3") + "/" + vDate_Start_D.ToString("MM/dd") +
                                                                                    "～" +
                                                                                    (vDate_End_D.Year - 1911).ToString("D3") + "/" + vDate_End_D.ToString("MM/dd"));
                                //wsExcel.AutoSizeColumn(0);
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 2));
                                wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("科目範圍：" + vSubjectNo);
                                //wsExcel.AutoSizeColumn(3);
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 3, 4));
                                vLinesNo++;
                                wsExcel.CreateRow(vLinesNo);
                                wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("科目代號：" + vSubjectNo);
                                //wsExcel.AutoSizeColumn(0);
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 1));
                                wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("科目名稱：" + vSubjectName);
                                //wsExcel.AutoSizeColumn(2);
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, 3));
                                wsExcel.GetRow(vLinesNo).CreateCell(4);
                                wsExcel.GetRow(vLinesNo).GetCell(4).SetCellValue("期初餘額：");
                                wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csRightAligment;
                                //wsExcel.AutoSizeColumn(4);
                                wsExcel.GetRow(vLinesNo).CreateCell(5);
                                wsExcel.GetRow(vLinesNo).GetCell(5).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(5).SetCellValue(vPBalance);
                                wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue(vPBalanceStr);
                                vLinesNo++;
                                wsExcel.CreateRow(vLinesNo);
                                wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("");
                                //wsExcel.AutoSizeColumn(0);
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 7));
                                //建立標題列
                                vLinesNo++;
                                wsExcel.CreateRow(vLinesNo);
                                wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("日期");
                                wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue("傳票號碼");
                                wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("部門");
                                wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("摘要");
                                wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("借方金額");
                                wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("貸方金額");
                                wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("餘");
                                wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("本期餘額");
                                wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csSubTitle;
                                wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("摘要二");
                                wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csSubTitle;
                            }
                        }
                    }
                    try
                    {
                        /*
                        MemoryStream msTarget = new MemoryStream();
                        wbExcel.Write(msTarget);
                        // 設定強制下載標頭
                        Response.AddHeader("Content-Disposition", string.Format("attachment; filename=科目分類帳轉出EXCEL.xls"));
                        // 輸出檔案
                        Response.BinaryWrite(msTarget.ToArray());
                        msTarget.Close();

                        Response.End(); //*/
                        string vFileName = "科目分類帳轉出EXCEL";
                        var msTarget = new NPOIMemoryStream();
                        msTarget.AllowClose = false;
                        wbExcel.Write(msTarget);
                        msTarget.Flush();
                        msTarget.Seek(0, SeekOrigin.Begin);
                        msTarget.AllowClose = true;

                        if (msTarget.Length > 0)
                        {
                            string vCalDateStr = ((eCalDate_Start_Search.Text.Trim() != "") && (eCalDate_End_Search.Text.Trim() != "")) ? "從 " + eCalDate_Start_Search.Text.Trim() + " 起至 " + eCalDate_End_Search.Text.Trim() + " 止" :
                                                 ((eCalDate_Start_Search.Text.Trim() != "") && (eCalDate_End_Search.Text.Trim() == "")) ? eCalDate_Start_Search.Text.Trim() :
                                                 ((eCalDate_Start_Search.Text.Trim() == "") && (eCalDate_End_Search.Text.Trim() != "")) ? eCalDate_End_Search.Text.Trim() : "";
                            string vSubjectNoStr = ((eSubjectNo_Start_Search.Text.Trim() != "") && (eSubjectNo_End_Search.Text.Trim() != "")) ? "從 " + eCalDate_Start_Search.Text.Trim() + " 起至 " + eCalDate_End_Search.Text.Trim() + " 止" :
                                                   ((eSubjectNo_Start_Search.Text.Trim() != "") && (eSubjectNo_End_Search.Text.Trim() == "")) ? eSubjectNo_Start_Search.Text.Trim() :
                                                   ((eSubjectNo_Start_Search.Text.Trim() == "") && (eSubjectNo_End_Search.Text.Trim() != "")) ? eSubjectNo_End_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_科目分類帳轉出EXCEL" + Environment.NewLine +
                                                 "TrialBalance.aspx" + Environment.NewLine +
                                                 "結算日期：" + vCalDateStr + Environment.NewLine +
                                                 "會計科目：" + vSubjectNoStr;
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
                        Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                        Response.Write("</" + "Script>");
                    }
                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 列印累進試算表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbTransToPrint_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelectStr();
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                //起始日期
                DateTime vDate_Start_D = (eCalDate_Start_Search.Text.Trim() != "") ? DateTime.Parse(eCalDate_Start_Search.Text.Trim()) : DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today, "C"));
                string vDate_Start = (eCalDate_Start_Search.Text.Trim() != "") ? PF.TransDateString(eCalDate_Start_Search.Text.Trim(), "B") : PF.GetMonthFirstDay(DateTime.Today, "B");
                //結束日期
                DateTime vDate_End_D = (eCalDate_End_Search.Text.Trim() != "") ? DateTime.Parse(eCalDate_End_Search.Text.Trim()) : DateTime.Parse(PF.GetMonthLastDay(DateTime.Today, "C"));
                string vDate_End = (eCalDate_End_Search.Text.Trim() != "") ? PF.TransDateString(eCalDate_End_Search.Text.Trim(), "B") : PF.GetMonthLastDay(DateTime.Today, "B");
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable();
                daPrint.Fill(dtPrint);

                ReportDataSource rdsPrint = new ReportDataSource("TrialBalanceP1", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\TrialBalanceP1.rdlc";
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "累進試算表"));
                rvPrint.LocalReport.SetParameters(new ReportParameter("StartDate", vDate_Start));
                rvPrint.LocalReport.SetParameters(new ReportParameter("EndDate", vDate_End));
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.Refresh();
                plReport.Visible = true;
                plSearch.Visible = false;
            }
        }

        /// <summary>
        /// 列印各科目分類帳明細
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbTransToPrint2_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            //string vSQLStr = "";
            string vSQLStr2 = "";
            //起始日期
            DateTime vDate_Start_D = (eCalDate_Start_Search.Text.Trim() != "") ? DateTime.Parse(eCalDate_Start_Search.Text.Trim()) : DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today, "C"));
            string vDate_Start = (eCalDate_Start_Search.Text.Trim() != "") ? PF.TransDateString(eCalDate_Start_Search.Text.Trim(), "B") : PF.GetMonthFirstDay(DateTime.Today, "B");
            //結束日期
            DateTime vDate_End_D = (eCalDate_End_Search.Text.Trim() != "") ? DateTime.Parse(eCalDate_End_Search.Text.Trim()) : DateTime.Parse(PF.GetMonthLastDay(DateTime.Today, "C"));
            string vDate_End = (eCalDate_End_Search.Text.Trim() != "") ? PF.TransDateString(eCalDate_End_Search.Text.Trim(), "B") : PF.GetMonthLastDay(DateTime.Today, "B");
            //前一年年底
            string vDate_LastYearLastDay = (DateTime.Parse(eCalDate_Start_Search.Text.Trim()).Year - 1).ToString("D4") + "/12/31";
            //下月首日
            string vDate_NextMonthFirstDay = (DateTime.Parse(eCalDate_End_Search.Text.Trim()).AddDays(1)).Year.ToString("D4") + "/" + (DateTime.Parse(eCalDate_End_Search.Text.Trim()).AddDays(1)).ToString("MM/dd");
            //當年度首日
            string vDate_ThisYearFirstDay = PF.GetYearFirstDay(DateTime.Parse(eCalDate_Start_Search.Text.Trim()), "B");
            //當月首日
            string vDate_ThisMonthFirst = PF.GetMonthFirstDay(DateTime.Parse(eCalDate_Start_Search.Text.Trim()), "B");
            /*
            vSQLStr = "select ta.[No], ta.REC_Date, ta.[Subject], ta.Memo, ta.[Name], ta.DepName, ta.Memo2, Debit, Credit, "+Environment.NewLine+
                      "       case when Left(ta.[Subject],1) in ('1', '5', '8', '9') or left(ta.[Subject], 2) = '62' then tb.SDebit - tb.SCredit else tb.SCredit - tb.SDebit end LastAmount "+Environment.NewLine+
                      "  from (" + Environment.NewLine +
                      "               select a.[No], a.REC_Date, a.[Subject], a.Meno as Memo, isnull(a.Debit, 0) as Debit, isnull(a.Credit, 0) as Credit, b.[Name], " + Environment.NewLine +
                      "                      (select[Name] from Department where DepNo = a.Department) DepName, " + Environment.NewLine +
                      "                      b.IsDebit as DebitType, a.Memo2 " + Environment.NewLine +
                      "                 from Account a, AC_Subject b " + Environment.NewLine +
                      "                where a.[Type] <> '6' and a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "                  and a.REC_Date >= '" + vDate_Start + "' and a.REC_Date < '" + vDate_NextMonthFirstDay + "' " + Environment.NewLine +
                      "                  and a.VoidDate IS NULL ) ta " + Environment.NewLine +
                      "               left join " + Environment.NewLine +
                      "              (select [Subject], sum(SCredit) SCredit , sum(SDebit) SDebit " + Environment.NewLine +
                      "                 from (select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                      "                         from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                      "                        where a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "                          and a.YYMM_DATE < '" + vDate_Start + "' and a.[Subject] < '4%' " + Environment.NewLine +
                      "                        group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                      "                        union all " + Environment.NewLine +
                      "                       select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                      "                         from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                      "                        where a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "                          and a.YYMM_DATE < '" + vDate_Start + "' and a.YYMM_DATE >= '" + vDate_ThisYearFirstDay + "' " + Environment.NewLine +
                      "                          and a.[Subject] > '4%' " + Environment.NewLine +
                      "                        group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                      "                        union all " + Environment.NewLine +
                      "                       select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                      "                         from Account a , AC_Subject b " + Environment.NewLine +
                      "                        where a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "                          and a.VoidDate IS NULL " + Environment.NewLine +
                      "                          and a.REC_Date >= '" + vDate_ThisMonthFirst + "' and a.REC_Date < '" + vDate_Start + "' " + Environment.NewLine +
                      "                        group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                      "                       ) t Group by[Subject] " + Environment.NewLine +
                      "              ) tb on tb.[SUBJECT] = ta.[Subject] " + Environment.NewLine +
                      " order by ta.Subject, ta.REC_Date "; //*/
            /*
            vSQLStr = "select t.[Subject], tb.[Name], --sum(SCredit) SCredit , sum(SDebit) SDebit, " + Environment.NewLine +
                      "       case when Left(t.[Subject],1) in ('1', '5', '8', '9') or left(t.[Subject], 2) = '62' then t.SDebit - t.SCredit else t.SCredit - t.SDebit end LastAmount, " + Environment.NewLine +
                      "       case when(Left(t.[Subject], 1) in ('1', '5', '8', '9') or left(t.[Subject], 2) = '62') and(t.SDebit - t.SCredit > 0) then '借' " + Environment.NewLine +
                      "            when(Left(t.[Subject], 1) in ('2', '3', '4', '7') and(t.SCredit - t.SDebit) < 0) then '借' else '貸' end DebitStr " + Environment.NewLine +
                      "  from (select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                      "          from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                      "         where a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "           and a.YYMM_DATE < '" + vDate_Start + "' and a.[Subject] < '4%' " + Environment.NewLine +
                      "         group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                      "         union all " + Environment.NewLine +
                      "        select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                      "          from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                      "         where a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "           and a.YYMM_DATE < '" + vDate_Start + "' and a.YYMM_DATE >= '" + vDate_ThisYearFirstDay + "' " + Environment.NewLine +
                      "           and a.[Subject] > '4%' " + Environment.NewLine +
                      "         group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                      "         union all " + Environment.NewLine +
                      "        select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                      "          from Account a , AC_Subject b " + Environment.NewLine +
                      "         where a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "           and a.VoidDate IS NULL " + Environment.NewLine +
                      "           and a.REC_Date >= '" + vDate_ThisMonthFirst + "' and a.REC_Date < '" + vDate_Start + "' " + Environment.NewLine +
                      "         group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                      "       ) t left join AC_Subject as tb on tb.[Subject] = t.[Subject] " + Environment.NewLine +
                      " Group by t.[Subject], tb.[Name], t.SDebit, t.SCredit"; //*/
            ///*
            vSQLStr2 = "select ta.[No], ta.REC_Date, ta.[Subject], ta.Memo, ta.[Name], ta.DepName, ta.Memo2, Debit, Credit, cast(0 as int) TempAmount, " + Environment.NewLine +
                       "       case when Left(ta.[Subject],1) in ('1', '5', '8', '9') or left(ta.[Subject], 2) = '62' then Debit - Credit else Credit - Debit end Amount, " + Environment.NewLine +
                       "       case when Left(ta.[Subject],1) in ('1', '5', '8', '9') or left(ta.[Subject], 2) = '62' then isnull(tb.SDebit, 0) - isnull(tb.SCredit, 0) else isnull(tb.SCredit, 0) - isnull(tb.SDebit, 0) end LastAmount, " + Environment.NewLine +
                       "       cast('' as varchar) AccountType, (Debit - Credit) as Balance " + Environment.NewLine +
                       "  from (" + Environment.NewLine +
                       "               select a.[No], a.REC_Date, a.[Subject], a.Meno as Memo, isnull(a.Debit, 0) as Debit, isnull(a.Credit, 0) as Credit, b.[Name], " + Environment.NewLine +
                       "                      (select[Name] from Department where DepNo = a.Department) DepName, " + Environment.NewLine +
                       "                      b.IsDebit as DebitType, a.Memo2 " + Environment.NewLine +
                       "                 from Account a, AC_Subject b " + Environment.NewLine +
                       "                where a.[Type] <> '6' and a.[Subject] = b.[Subject] " + Environment.NewLine +
                       "                  and a.REC_Date >= '" + vDate_Start + "' and a.REC_Date < '" + vDate_NextMonthFirstDay + "' " + Environment.NewLine +
                       "                  and a.VoidDate IS NULL ) ta " + Environment.NewLine +
                       "               left join " + Environment.NewLine +
                       "              (select [Subject], sum(SCredit) SCredit , sum(SDebit) SDebit " + Environment.NewLine +
                       "                 from (select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                       "                         from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                       "                        where a.[Subject] = b.[Subject] " + Environment.NewLine +
                       "                          and a.YYMM_DATE < '" + vDate_Start + "' and a.[Subject] < '4%' " + Environment.NewLine +
                       "                        group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                       "                        union all " + Environment.NewLine +
                       "                       select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                       "                         from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                       "                        where a.[Subject] = b.[Subject] " + Environment.NewLine +
                       "                          and a.YYMM_DATE < '" + vDate_Start + "' and a.YYMM_DATE >= '" + vDate_ThisYearFirstDay + "' " + Environment.NewLine +
                       "                          and a.[Subject] > '4%' " + Environment.NewLine +
                       "                        group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                       "                        union all " + Environment.NewLine +
                       "                       select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                       "                         from Account a , AC_Subject b " + Environment.NewLine +
                       "                        where a.[Subject] = b.[Subject] " + Environment.NewLine +
                       "                          and a.VoidDate IS NULL " + Environment.NewLine +
                       "                          and a.REC_Date >= '" + vDate_ThisMonthFirst + "' and a.REC_Date < '" + vDate_Start + "' " + Environment.NewLine +
                       "                        group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                       "                       ) t Group by[Subject] " + Environment.NewLine +
                       "              ) tb on tb.[SUBJECT] = ta.[Subject] " + Environment.NewLine +
                       " order by ta.Subject, ta.REC_Date, ta.[No] "; //*/

            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                //SqlDataAdapter daPrint = new SqlDataAdapter(vSQLStr, connPrint);
                SqlDataAdapter daPrint_Sub = new SqlDataAdapter(vSQLStr2, connPrint);
                connPrint.Open();
                //DataTable dtPrint = new DataTable();
                //daPrint.Fill(dtPrint);
                DataTable dtPrint_Sub = new DataTable();
                daPrint_Sub.Fill(dtPrint_Sub);
                if (dtPrint_Sub.Rows.Count > 0)
                {
                    string vSubject_Temp = "";
                    string vAccountType_Temp = "";
                    int vTAmount_Temp = 0;
                    int vTBalance_Temp = 0;
                    int vLastAmount_Temp = 0;
                    int vTempINT = 0;
                    for (int i = 0; i < dtPrint_Sub.Rows.Count; i++)
                    {
                        DataRow tempRow = dtPrint_Sub.Rows[i];
                        if (dtPrint_Sub.Rows[i]["Subject"].ToString().Trim() != vSubject_Temp)
                        {
                            vSubject_Temp = dtPrint_Sub.Rows[i]["Subject"].ToString().Trim();
                            vTAmount_Temp = 0;
                        }
                        vTBalance_Temp += Int32.Parse(dtPrint_Sub.Rows[i]["Balance"].ToString().Trim());
                        vLastAmount_Temp = Int32.Parse(dtPrint_Sub.Rows[i]["LastAmount"].ToString().Trim());
                        if ((((vSubject_Temp.Substring(0, 1) == "1") || (vSubject_Temp.Substring(0, 1) == "5") || (vSubject_Temp.Substring(0, 1) == "8") || (vSubject_Temp.Substring(0, 1) == "9") || (vSubject_Temp.Substring(0, 2) == "62")) && (vLastAmount_Temp > 0)) ||
                            (((vSubject_Temp.Substring(0, 1) == "2") || (vSubject_Temp.Substring(0, 1) == "3") || (vSubject_Temp.Substring(0, 1) == "4") || (vSubject_Temp.Substring(0, 1) == "7")) && (vLastAmount_Temp < 0)))
                        {
                            //借方會科
                            vTempINT = vTBalance_Temp + vLastAmount_Temp;
                        }
                        else
                        {
                            //貸方會科
                            vTempINT = vTBalance_Temp - vLastAmount_Temp;
                        }
                        vAccountType_Temp = (vTempINT < 0) ? "貸" : "借";
                        dtPrint_Sub.Rows[i]["AccountType"] = vAccountType_Temp;
                        vTAmount_Temp += Int32.Parse(dtPrint_Sub.Rows[i]["Amount"].ToString().Trim());
                        dtPrint_Sub.Rows[i]["TempAmount"] = vTAmount_Temp;
                    }
                }

                //ReportDataSource rdsPrint = new ReportDataSource("TrialBalanceP2Main", dtPrint);
                ReportDataSource rdsPrint = new ReportDataSource("TrialBalanceP2Sub", dtPrint_Sub);
                //ReportDataSource rdsPrint_Sub = new ReportDataSource("TrialBalanceP2Sub", dtPrint_Sub);

                //rvPrint.LocalReport.SubreportProcessing += new SubreportProcessingEventHandler(LocalReport_SubreportProcessing);
                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\TrialBalanceP2.rdlc";
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "科目分類帳明細表"));
                rvPrint.LocalReport.SetParameters(new ReportParameter("StartDate", vDate_Start));
                rvPrint.LocalReport.SetParameters(new ReportParameter("EndDate", vDate_End));
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                //rvPrint.LocalReport.DataSources.Add(rdsPrint_Sub);
                rvPrint.LocalReport.Refresh();
                //plReport2.Visible = true;
                plReport.Visible = true;
                plSearch.Visible = false;
            }
        }

        private void LocalReport_SubreportProcessing(object sender, SubreportProcessingEventArgs e)
        {
            //e.DataSources.Add(new ReportDataSource("TrialBalanceP2Sub", GetSubReportData()));
        }

        /// <summary>
        /// 不帶參數的做法
        /// </summary>
        /// <returns></returns>
        private DataTable GetSubReportData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DataTable dtResult = new DataTable();
            string vSQLStr = "";
            //起始日期
            DateTime vDate_Start_D = (eCalDate_Start_Search.Text.Trim() != "") ? DateTime.Parse(eCalDate_Start_Search.Text.Trim()) : DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today, "C"));
            string vDate_Start = (eCalDate_Start_Search.Text.Trim() != "") ? PF.TransDateString(eCalDate_Start_Search.Text.Trim(), "B") : PF.GetMonthFirstDay(DateTime.Today, "B");
            //結束日期
            DateTime vDate_End_D = (eCalDate_End_Search.Text.Trim() != "") ? DateTime.Parse(eCalDate_End_Search.Text.Trim()) : DateTime.Parse(PF.GetMonthLastDay(DateTime.Today, "C"));
            string vDate_End = (eCalDate_End_Search.Text.Trim() != "") ? PF.TransDateString(eCalDate_End_Search.Text.Trim(), "B") : PF.GetMonthLastDay(DateTime.Today, "B");
            //前一年年底
            string vDate_LastYearLastDay = (DateTime.Parse(eCalDate_Start_Search.Text.Trim()).Year - 1).ToString("D4") + "/12/31";
            //下月首日
            string vDate_NextMonthFirstDay = (DateTime.Parse(eCalDate_End_Search.Text.Trim()).AddDays(1)).Year.ToString("D4") + "/" + (DateTime.Parse(eCalDate_End_Search.Text.Trim()).AddDays(1)).ToString("MM/dd");
            //當年度首日
            string vDate_ThisYearFirstDay = PF.GetYearFirstDay(DateTime.Parse(eCalDate_Start_Search.Text.Trim()), "B");
            //當月首日
            string vDate_ThisMonthFirst = PF.GetMonthFirstDay(DateTime.Parse(eCalDate_Start_Search.Text.Trim()), "B");
            vSQLStr = "select ta.[No], ta.REC_Date, ta.[Subject], ta.Memo, ta.[Name], ta.DepName, ta.Memo2, Debit, Credit, " + Environment.NewLine +
                      "       case when Left(ta.[Subject],1) in ('1', '5', '8', '9') or left(ta.[Subject], 2) = '62' then tb.SDebit - tb.SCredit else tb.SCredit - tb.SDebit end LastAmount " + Environment.NewLine +
                      "  from (" + Environment.NewLine +
                      "               select a.[No], a.REC_Date, a.[Subject], a.Meno as Memo, isnull(a.Debit, 0) as Debit, isnull(a.Credit, 0) as Credit, b.[Name], " + Environment.NewLine +
                      "                      (select[Name] from Department where DepNo = a.Department) DepName, " + Environment.NewLine +
                      "                      b.IsDebit as DebitType, a.Memo2 " + Environment.NewLine +
                      "                 from Account a, AC_Subject b " + Environment.NewLine +
                      "                where a.[Type] <> '6' and a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "                  and a.REC_Date >= '" + vDate_Start + "' and a.REC_Date < '" + vDate_NextMonthFirstDay + "' " + Environment.NewLine +
                      "                  and a.VoidDate IS NULL ) ta " + Environment.NewLine +
                      "               left join " + Environment.NewLine +
                      "              (select [Subject], sum(SCredit) SCredit , sum(SDebit) SDebit " + Environment.NewLine +
                      "                 from (select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                      "                         from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                      "                        where a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "                          and a.YYMM_DATE < '" + vDate_Start + "' and a.[Subject] < '4%' " + Environment.NewLine +
                      "                        group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                      "                        union all " + Environment.NewLine +
                      "                       select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                      "                         from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                      "                        where a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "                          and a.YYMM_DATE < '" + vDate_Start + "' and a.YYMM_DATE >= '" + vDate_ThisYearFirstDay + "' " + Environment.NewLine +
                      "                          and a.[Subject] > '4%' " + Environment.NewLine +
                      "                        group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                      "                        union all " + Environment.NewLine +
                      "                       select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                      "                         from Account a , AC_Subject b " + Environment.NewLine +
                      "                        where a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "                          and a.VoidDate IS NULL " + Environment.NewLine +
                      "                          and a.REC_Date >= '" + vDate_ThisMonthFirst + "' and a.REC_Date < '" + vDate_Start + "' " + Environment.NewLine +
                      "                        group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                      "                       ) t Group by[Subject] " + Environment.NewLine +
                      "              ) tb on tb.[SUBJECT] = ta.[Subject] " + Environment.NewLine +
                      " order by ta.[Subject], ta.REC_Date ";

            using (SqlConnection connSubReport = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daSubReport = new SqlDataAdapter(vSQLStr, connSubReport);
                connSubReport.Open();
                daSubReport.Fill(dtResult);
            }
            return dtResult;
        }

        /// <summary>
        /// 帶參數的做法
        /// </summary>
        /// <param name="fSubjectID"></param>
        /// <returns></returns>
        private DataTable GetSubReportData(string fSubjectID)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DataTable dtResult = new DataTable();
            string vSQLStr = "";
            //起始日期
            DateTime vDate_Start_D = (eCalDate_Start_Search.Text.Trim() != "") ? DateTime.Parse(eCalDate_Start_Search.Text.Trim()) : DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today, "C"));
            string vDate_Start = (eCalDate_Start_Search.Text.Trim() != "") ? PF.TransDateString(eCalDate_Start_Search.Text.Trim(), "B") : PF.GetMonthFirstDay(DateTime.Today, "B");
            //結束日期
            DateTime vDate_End_D = (eCalDate_End_Search.Text.Trim() != "") ? DateTime.Parse(eCalDate_End_Search.Text.Trim()) : DateTime.Parse(PF.GetMonthLastDay(DateTime.Today, "C"));
            string vDate_End = (eCalDate_End_Search.Text.Trim() != "") ? PF.TransDateString(eCalDate_End_Search.Text.Trim(), "B") : PF.GetMonthLastDay(DateTime.Today, "B");
            //前一年年底
            string vDate_LastYearLastDay = (DateTime.Parse(eCalDate_Start_Search.Text.Trim()).Year - 1).ToString("D4") + "/12/31";
            //下月首日
            string vDate_NextMonthFirstDay = (DateTime.Parse(eCalDate_End_Search.Text.Trim()).AddDays(1)).Year.ToString("D4") + "/" + (DateTime.Parse(eCalDate_End_Search.Text.Trim()).AddDays(1)).ToString("MM/dd");
            //當年度首日
            string vDate_ThisYearFirstDay = PF.GetYearFirstDay(DateTime.Parse(eCalDate_Start_Search.Text.Trim()), "B");
            //當月首日
            string vDate_ThisMonthFirst = PF.GetMonthFirstDay(DateTime.Parse(eCalDate_Start_Search.Text.Trim()), "B");
            vSQLStr = "select ta.[No], ta.REC_Date, ta.[Subject], ta.Memo, ta.[Name], ta.DepName, ta.Memo2, Debit, Credit, " + Environment.NewLine +
                      "       case when Left(ta.[Subject],1) in ('1', '5', '8', '9') or left(ta.[Subject], 2) = '62' then tb.SDebit - tb.SCredit else tb.SCredit - tb.SDebit end LastAmount " + Environment.NewLine +
                      "  from (" + Environment.NewLine +
                      "               select a.[No], a.REC_Date, a.[Subject], a.Meno as Memo, isnull(a.Debit, 0) as Debit, isnull(a.Credit, 0) as Credit, b.[Name], " + Environment.NewLine +
                      "                      (select[Name] from Department where DepNo = a.Department) DepName, " + Environment.NewLine +
                      "                      b.IsDebit as DebitType, a.Memo2 " + Environment.NewLine +
                      "                 from Account a, AC_Subject b " + Environment.NewLine +
                      "                where a.[Type] <> '6' and a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "                  and a.REC_Date >= '" + vDate_Start + "' and a.REC_Date < '" + vDate_NextMonthFirstDay + "' " + Environment.NewLine +
                      "                  and a.VoidDate IS NULL ) ta " + Environment.NewLine +
                      "               left join " + Environment.NewLine +
                      "              (select [Subject], sum(SCredit) SCredit , sum(SDebit) SDebit " + Environment.NewLine +
                      "                 from (select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                      "                         from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                      "                        where a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "                          and a.YYMM_DATE < '" + vDate_Start + "' and a.[Subject] < '4%' " + Environment.NewLine +
                      "                        group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                      "                        union all " + Environment.NewLine +
                      "                       select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                      "                         from TOTAL_Account a, AC_Subject b " + Environment.NewLine +
                      "                        where a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "                          and a.YYMM_DATE < '" + vDate_Start + "' and a.YYMM_DATE >= '" + vDate_ThisYearFirstDay + "' " + Environment.NewLine +
                      "                          and a.[Subject] > '4%' " + Environment.NewLine +
                      "                        group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                      "                        union all " + Environment.NewLine +
                      "                       select a.[Subject], sum(isnull(a.Debit, 0)) as SDebit, sum(isnull(a.Credit, 0)) as SCredit " + Environment.NewLine +
                      "                         from Account a , AC_Subject b " + Environment.NewLine +
                      "                        where a.[Subject] = b.[Subject] " + Environment.NewLine +
                      "                          and a.VoidDate IS NULL " + Environment.NewLine +
                      "                          and a.REC_Date >= '" + vDate_ThisMonthFirst + "' and a.REC_Date < '" + vDate_Start + "' " + Environment.NewLine +
                      "                        group by a.[Subject], b.[Name], b.IsDebit " + Environment.NewLine +
                      "                       ) t Group by[Subject] " + Environment.NewLine +
                      "              ) tb on tb.[SUBJECT] = ta.[Subject] " + Environment.NewLine +
                      " where ta.[Subject] = '" + fSubjectID + "' " + Environment.NewLine +
                      " order by ta.[Subject], ta.REC_Date ";

            using (SqlConnection connSubReport = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daSubReport = new SqlDataAdapter(vSQLStr, connSubReport);
                connSubReport.Open();
                daSubReport.Fill(dtResult);
            }
            return dtResult;
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            rvPrint.LocalReport.DataSources.Clear();
            plReport.Visible = false;
            //plReport2.Visible = false;
            plSearch.Visible = true;
        }

        /*
        protected void rvPrint2_Drillthrough(object sender, DrillthroughEventArgs e)
        {
            string vSubjectIS = "";

            LocalReport vSubReport = (LocalReport)e.Report;
            IList<ReportParameter> vList = vSubReport.OriginalParametersToDrillthrough;
            foreach (ReportParameter vParam in vList)
            {
                vSubjectIS = vParam.Values[0].ToString().Trim();
            }
            vSubReport.DataSources.Add(new ReportDataSource("TrialBalanceP2Sub", GetSubReportData(vSubjectIS)));
        } //*/
    }
}