using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class CalAvgPayment : System.Web.UI.Page
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

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbOK_Click(object sender, EventArgs e)
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

            DateTime vCalDate = DateTime.Today;

            //先決定要抓哪幾個月的資料
            //第一個月
            int vYear_Start = (((vCalDate.Month) - 3) <= 0) ? vCalDate.Year - 1 : vCalDate.Year;
            int vYear_Start_C = vYear_Start - 1911;
            int vMonth_Start = (((vCalDate.Month) - 3) <= 0) ? vCalDate.Month + 9 : vCalDate.Month - 3;
            DateTime vStartDate = DateTime.Parse(vYear_Start.ToString("D4") + "/" + vMonth_Start.ToString("D2") + "/01");
            string vStartDate_C = vYear_Start_C.ToString("D3") + "年" + vMonth_Start.ToString("D2") + "月";
            //第二個月
            int vYear_Mid = (((vCalDate.Month) - 2) <= 0) ? vCalDate.Year - 1 : vCalDate.Year;
            int vYear_Mid_C = vYear_Mid - 1911;
            int vMonth_Mid = (((vCalDate.Month) - 2) <= 0) ? vCalDate.Month + 10 : vCalDate.Month - 2;
            DateTime vMidDate = DateTime.Parse(vYear_Mid.ToString("D4") + "/" + vMonth_Mid.ToString("D2") + "/01");
            string vMidDate_C = vYear_Mid_C.ToString("D3") + "年" + vMonth_Mid.ToString("D2") + "月";
            //第三個月
            int vYear_End = (((vCalDate.Month) - 1) <= 0) ? vCalDate.Year - 1 : vCalDate.Year;
            int vYear_End_C = vYear_End - 1911;
            int vMonth_End = (((vCalDate.Month) - 1) <= 0) ? vCalDate.Month + 11 : vCalDate.Month - 1;
            DateTime vEndDate = DateTime.Parse(vYear_End.ToString("D4") + "/" + vMonth_End.ToString("D2") + "/01");
            string vEndDate_C = vYear_End_C.ToString("D3") + "年" + vMonth_End.ToString("D2") + "月";

            string vSelectSQLStr = "";
            int vLinesNo = 0;
            string vHeaderText = "";
            string vCellName = "";
            //取得 SQL 主機的連結字串
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            SqlConnection connAvgPayment = new SqlConnection(vConnStr);
            SqlCommand cmdAvgPayment;
            SqlDataReader drAvgPayment;

            vSelectSQLStr = "select LiClass, isnull(LiAMT, 0) LiAMT, cast((isnull(LiAMT, 0) * 10.5 * 0.01 * 0.2) as decimal(10,0)) LiFee " +
                            "  from Labor order by LiAMT ";
            cmdAvgPayment = new SqlCommand(vSelectSQLStr, connAvgPayment);
            connAvgPayment.Open();
            drAvgPayment = cmdAvgPayment.ExecuteReader();
            if (drAvgPayment.HasRows)
            {
                vHeaderText = "";
                //新增一個工作表
                wsExcel = (HSSFSheet)wbExcel.CreateSheet("勞保級距");
                //寫入標題列
                vLinesNo = 0;
                wsExcel.CreateRow(vLinesNo);
                for (int i = 0; i < drAvgPayment.FieldCount; i++)
                {
                    vHeaderText = (drAvgPayment.GetName(i) == "LiClass") ? "級距編號" :
                                  (drAvgPayment.GetName(i) == "LiAMT") ? "投保金額" :
                                  (drAvgPayment.GetName(i) == "LiFee") ? "勞保金額" : drAvgPayment.GetName(i);
                    wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                }
                vLinesNo++;
                while (drAvgPayment.Read())
                {
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drAvgPayment.FieldCount; i++)
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(i);
                        if ((drAvgPayment.GetName(i) == "LiAMT") || (drAvgPayment.GetName(i) == "LiFee"))
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drAvgPayment[i].ToString()));
                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                        }
                        else
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drAvgPayment[i].ToString());
                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                        }
                    }
                    vLinesNo++;
                }
            }
            connAvgPayment.Close();
            cmdAvgPayment.Dispose();
            drAvgPayment.Close();

            vSelectSQLStr = "select HiClass, isnull(HiAMT, 0) HiAMT, cast((isnull(HiAMT, 0) * 4.69 * 0.01 * 0.3) as decimal(10,0)) HiFee " +
                            "  from Health order by HiAMT ";
            cmdAvgPayment = new SqlCommand(vSelectSQLStr, connAvgPayment);
            connAvgPayment.Open();
            drAvgPayment = cmdAvgPayment.ExecuteReader();
            if (drAvgPayment.HasRows)
            {
                vHeaderText = "";
                //新增一個工作表
                wsExcel = (HSSFSheet)wbExcel.CreateSheet("健保級距");
                //寫入標題列
                vLinesNo = 0;
                wsExcel.CreateRow(vLinesNo);
                for (int i = 0; i < drAvgPayment.FieldCount; i++)
                {
                    vHeaderText = (drAvgPayment.GetName(i) == "HiClass") ? "級距編號" :
                                  (drAvgPayment.GetName(i) == "HiAMT") ? "投保金額" :
                                  (drAvgPayment.GetName(i) == "HiFee") ? "健保金額" : drAvgPayment.GetName(i);
                    wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                }
                vLinesNo++;
                while (drAvgPayment.Read())
                {
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drAvgPayment.FieldCount; i++)
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(i);
                        if ((drAvgPayment.GetName(i) == "HiAMT") || (drAvgPayment.GetName(i) == "HiFee"))
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drAvgPayment[i].ToString()));
                        }
                        else
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drAvgPayment[i].ToString());
                        }
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                    }
                    vLinesNo++;
                }
            }
            connAvgPayment.Close();
            cmdAvgPayment.Dispose();
            drAvgPayment.Close();

            vSelectSQLStr = "select e.DepNo, e.Title, " + Environment.NewLine +
                            "       (select CLassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        TITLE' and ClassNo = e.Title) Title_C, " + Environment.NewLine +
                            "       t.EmpNo, e.[Name], e.Assumeday, " + Environment.NewLine +
                            "       t.YM_01_T as [" + vStartDate_C + "], t.ClassBounds_01_T as [" + vStartDate_C + "授課], " + Environment.NewLine +
                            "       t.YM_02_T as [" + vMidDate_C + "], t.ClassBounds_02_T as [" + vMidDate_C + "授課], " + Environment.NewLine +
                            "       t.YM_03_T as [" + vEndDate_C + "], t.ClassBounds_03_T as [" + vEndDate_C + "授課], " + Environment.NewLine +
                            "       t.ClassBounds, cast((t.ClassBounds / 3) as decimal(10, 0)) ClassBounds_AVG, " + Environment.NewLine +
                            "       cast((t.YM_Total / 3) as decimal(10, 0)) PaymentAVG, cast(((t.Payment_New) / 3) as decimal(10, 0)) PaymentAVG_New, e.IDCardNo, e.Birthday, " + Environment.NewLine +
                            //2023.08.09 把勞健保級距資料表裡面的 BLOCK 欄位拿來設定該級距我司有沒有使用，我司實務上有使用的級距設成 'Y'的級距在這個表單中才可以取到資料
                            "       e.LaiAMT, isnull((select MIN(LiAMT) from LABOR where LiAMT >= cast(((t.Payment_New + t.ClassBounds) / 3) as decimal(10, 0)) and BLOCK = 'Y' ), (select MAX(LiAMT) from LABOR)) LiAMT_New, " + Environment.NewLine +
                            "       (e.LaiAMT - isnull((select MIN(LiAMT) from LABOR where LiAMT >= cast(((t.Payment_New + t.ClassBounds) / 3) as decimal(10, 0)) and BLOCK = 'Y' ), (select MAX(LiAMT) from LABOR))) LiAMT_Diff, " + Environment.NewLine +
                            "       e.HiAMT, isnull((select MIN(HiAMT) from Health where HiAMT >= cast(((t.Payment_New + t.ClassBounds) / 3) as decimal(10, 0)) and BLOCK = 'Y' ), (select MAX(HiAMT) from Health)) HiAMT_New, " + Environment.NewLine +
                            "       (e.HiAMT - isnull((select MIN(HiAMT) from Health where HiAMT >= cast(((t.Payment_New + t.ClassBounds) / 3) as decimal(10, 0)) and BLOCK = 'Y' ), (select MAX(HiAMT) from Health))) HiAMT_Diff, " + Environment.NewLine +
                            "       t.LiFee, t.HiFee, " + Environment.NewLine +
                            "       case when e.RetireType = '舊制' then e.RetireType else '' end RetireType " + Environment.NewLine +
                            "  from ( " + Environment.NewLine +
                            "        select z.EmpNo, sum(z.YM_01) YM_01_T, sum(z.YM_02) YM_02_T, sum(z.YM_03) YM_03_T, sum(z.YM_01 + z.YM_02 + z.YM_03) YM_Total, " + Environment.NewLine +
                            "               sum(z.LiFee) LiFee, sum(z.HiFee) HiFee, " + Environment.NewLine +
                            "               sum(z.ClassBounds_01) ClassBounds_01_T, sum(z.ClassBounds_02) ClassBounds_02_T, sum(z.ClassBounds_03) ClassBounds_03_T, " + Environment.NewLine +
                            "               sum(z.ClassBounds_01 + z.ClassBounds_02 + z.ClassBounds_03) ClassBounds, " + Environment.NewLine +
                            "               sum(z.YM_01 + z.YM_02 + z.YM_03 - z.ClassBounds_01 - z.ClassBounds_02 - z.ClassBounds_03) Payment_New " + Environment.NewLine +
                            "          from ( " + Environment.NewLine +
                            "                select EmpNo, GivCash YM_01, cast(0 as float) YM_02, cast(0 as float) YM_03, " + Environment.NewLine +
                            "                       cast(0 as float) LiFee, cast(0 as float) HiFee, " + Environment.NewLine +
                            "                       isnull((select sum(Expense) from MSHZ where PayDate = '" + vStartDate.ToString("yyyy/MM/dd") + "' and MSHZ.EmpNo = PayRec.EmpNo and PayBNo = '1413'), 0) ClassBounds_01, " + Environment.NewLine +
                            "                       cast(0 as float) ClassBounds_02, cast(0 as float) ClassBounds_03 " + Environment.NewLine +
                            "                  from PayRec " + Environment.NewLine +
                            "                 where PayDate = '" + vStartDate.ToString("yyyy/MM/dd") + "' and PayDur = '1' and DepNo > '00' " + Environment.NewLine +
                            "                 union all " + Environment.NewLine +
                            "                select EmpNo, cast(0 as float) YM_01, GivCash YM_02, cast(0 as float) YM_03, " + Environment.NewLine +
                            "                       cast(0 as float) LiFee, cast(0 as float) HiFee, cast(0 as float) Class_01, " + Environment.NewLine +
                            "                       isnull((select sum(Expense) from MSHZ where PayDate = '" + vMidDate.ToString("yyyy/MM/dd") + "' and MSHZ.EmpNo = PayRec.EmpNo and PayBNo = '1413'), 0) ClassBounds_02, " + Environment.NewLine +
                            "                       cast(0 as float) ClassBount_03 " + Environment.NewLine +
                            "                  from PayRec " + Environment.NewLine +
                            "                 where PayDate = '" + vMidDate.ToString("yyyy/MM/dd") + "' and PayDur = '1' and DepNo > '00' " + Environment.NewLine +
                            "                 union all " + Environment.NewLine +
                            "                select EmpNo, cast(0 as float) YM_01, cast(0 as float) YM_02, GivCash YM_03, LiFee, HiFee, cast(0 as float) ClassBounds_01, cast(0 as float) CLassBounds_02, " + Environment.NewLine +
                            "                       isnull((select sum(Expense) from MSHZ where PayDate = '" + vEndDate.ToString("yyyy/MM/dd") + "' and MSHZ.EmpNo = PayRec.EmpNo and PayBNo = '1413'), 0) ClassBounds_03 " + Environment.NewLine +
                            "                  from PayRec " + Environment.NewLine +
                            "                 where PayDate = '" + vEndDate.ToString("yyyy/MM/dd") + "' and PayDur = '1' and DepNo > '00' " + Environment.NewLine +
                            "               ) z " + Environment.NewLine +
                            "         group by z.EmpNo " + Environment.NewLine +
                            "       ) t left join Employee e on e.EmpNo = t.EmpNo " + Environment.NewLine +
                            " order by DepNo, Title, EmpNo ";
            cmdAvgPayment = new SqlCommand(vSelectSQLStr, connAvgPayment);
            connAvgPayment.Open();
            drAvgPayment = cmdAvgPayment.ExecuteReader();
            if (drAvgPayment.HasRows)
            {
                vHeaderText = "";
                //新增一個工作表
                wsExcel = (HSSFSheet)wbExcel.CreateSheet("平均工資（原始）");
                //寫入標題列
                vLinesNo = 0;
                vCellName = "";
                wsExcel.CreateRow(vLinesNo);
                for (int i = 0; i < drAvgPayment.FieldCount; i++)
                {
                    vHeaderText = (drAvgPayment.GetName(i) == "DepNo") ? "單位編號" :
                                  (drAvgPayment.GetName(i) == "Title") ? "職稱編號" :
                                  (drAvgPayment.GetName(i) == "Title_C") ? "職稱" :
                                  (drAvgPayment.GetName(i) == "EmpNo") ? "員工工號" :
                                  (drAvgPayment.GetName(i) == "Name") ? "姓名" :
                                  (drAvgPayment.GetName(i) == "Assumeday") ? "到職日" :
                                  (drAvgPayment.GetName(i) == "ClassBounds") ? "授課津貼" :
                                  (drAvgPayment.GetName(i) == "ClassBounds_AVG") ? "平均授課津貼" :
                                  (drAvgPayment.GetName(i) == "PaymentAVG") ? "平均工資" : //7
                                  (drAvgPayment.GetName(i) == "PaymentAVG_New") ? "平均工資(新)" :
                                  (drAvgPayment.GetName(i) == "IDCardNo") ? "身分證字號" :
                                  (drAvgPayment.GetName(i) == "Birthday") ? "出生日期" :
                                  (drAvgPayment.GetName(i) == "LiAMT_New") ? "新勞保對應費用" :
                                  (drAvgPayment.GetName(i) == "HiAMT_New") ? "新健保對應費用" : //7
                                  (drAvgPayment.GetName(i) == "LaiAMT") ? "勞保級距" :
                                  (drAvgPayment.GetName(i) == "LiFee") ? "勞保費" :
                                  (drAvgPayment.GetName(i) == "HiAMT") ? "健保級距" : //7
                                  (drAvgPayment.GetName(i) == "HiFee") ? "健保費" : 
                                  (drAvgPayment.GetName(i) == "LiAMT_Diff") ? "勞保級距差" :
                                  (drAvgPayment.GetName(i) == "HiAMT_Diff") ? "健保級距差" :
                                  (drAvgPayment.GetName(i) == "RetireType") ? "新舊制" : drAvgPayment.GetName(i);
                    wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                }
                vLinesNo++;
                while (drAvgPayment.Read())
                {
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drAvgPayment.FieldCount; i++)
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(i);
                        vCellName = drAvgPayment.GetName(i);
                        if ((vCellName == "ClassBounds") ||
                            (vCellName == "ClassBounds_AVG") ||
                            //(vCellName == "PaymentAVG") ||
                            (vCellName == "PaymentAVG_New") ||
                            (vCellName == "LiAMT_New") ||
                            //(vCellName == "HiAMT_New") ||
                            (vCellName == "LaiAMT") ||
                            (vCellName == "LiFee") ||
                            //(vCellName == "HiAMT") ||
                            (vCellName == "HiFee") ||
                            (vCellName == "LiAMT_Diff") ||
                            (vCellName == "HiAMT_Diff") ||
                            (vCellName == vStartDate_C + "授課") ||
                            (vCellName == vMidDate_C + "授課") ||
                            (vCellName == vEndDate_C + "授課") ||
                            (vCellName == vStartDate_C) ||
                            (vCellName == vMidDate_C) ||
                            (vCellName == vEndDate_C))
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drAvgPayment[i].ToString()));
                        }
                        else if ((vCellName == "PaymentAVG") ||
                                 (vCellName == "HiAMT_New")||
                                 (vCellName == "HiAMT"))
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(Int32.Parse(drAvgPayment[i].ToString()).ToString("D6"));
                        }
                        else if (vCellName == "Birthday")
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue((DateTime.Parse(drAvgPayment[i].ToString()).Year - 1911).ToString("D3") +
                                                                             DateTime.Parse(drAvgPayment[i].ToString()).Month.ToString("D2") +
                                                                             DateTime.Parse(drAvgPayment[i].ToString()).Day.ToString("D2"));
                        }
                        else if (vCellName == "Assumeday")
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue((DateTime.Parse(drAvgPayment[i].ToString()).Year - 1911).ToString("D3") + "/" +
                                                                             DateTime.Parse(drAvgPayment[i].ToString()).Month.ToString("D2") + "/" +
                                                                             DateTime.Parse(drAvgPayment[i].ToString()).Day.ToString("D2"));
                        }
                        else
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drAvgPayment[i].ToString());
                        }
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                    }
                    vLinesNo++;
                }
            }
            connAvgPayment.Close();
            cmdAvgPayment.Dispose();
            drAvgPayment.Close();

            //先抓指定範圍的薪資資料
            vSelectSQLStr = "select (cast((year(PayDate) - 1911) as varchar) + '年' + cast(month(PayDate) as varchar) + '月') PayDate, DepNo, Title, " +
                            "       (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        TITLE' and ClassNo = Payrec.Title) Title_C, " +
                            "       EmpNo, [Name], GivCash, LaiAMT, LiFee, HiAMT, HiFee " +
                            "  from PayREC " +
                            " where PayDate between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "' " +
                            "   and PayDur = '1' " +
                            " order by PayDate, DepNo, Title, EmpNo";
            cmdAvgPayment = new SqlCommand(vSelectSQLStr, connAvgPayment);
            connAvgPayment.Open();
            drAvgPayment = cmdAvgPayment.ExecuteReader();

            if (drAvgPayment.HasRows)
            {
                string vPayDate = "";
                vLinesNo = 0;
                vHeaderText = "";
                vCellName = "";

                while (drAvgPayment.Read())
                {
                    if (drAvgPayment["PayDate"].ToString() != vPayDate)
                    {
                        vLinesNo = 0;
                        vPayDate = drAvgPayment["PayDate"].ToString();
                        wsExcel = (HSSFSheet)wbExcel.CreateSheet(vPayDate);
                        //寫入標題列
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 1; i < drAvgPayment.FieldCount; i++)
                        {
                            vHeaderText = (drAvgPayment.GetName(i) == "DepNo") ? "單位編號" :
                                          (drAvgPayment.GetName(i) == "Title") ? "職稱編號" :
                                          (drAvgPayment.GetName(i) == "Title_C") ? "職稱" :
                                          (drAvgPayment.GetName(i) == "EmpNo") ? "員工工號" :
                                          (drAvgPayment.GetName(i) == "Name") ? "姓名" :
                                          (drAvgPayment.GetName(i) == "GivCash") ? "應發金額" :
                                          (drAvgPayment.GetName(i) == "LaiAMT") ? "勞保保額" :
                                          (drAvgPayment.GetName(i) == "LiFee") ? "勞保金額" :
                                          (drAvgPayment.GetName(i) == "HiAMT") ? "健保保額" :
                                          (drAvgPayment.GetName(i) == "HiFee") ? "健保金額" : drAvgPayment.GetName(i);
                            wsExcel.GetRow(vLinesNo).CreateCell(i - 1).SetCellValue(vHeaderText);
                            wsExcel.GetRow(vLinesNo).GetCell(i - 1).CellStyle = csTitle;
                        }
                        vLinesNo++;
                    }
                    wsExcel = (HSSFSheet)wbExcel.GetSheet(vPayDate);
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 1; i < drAvgPayment.FieldCount; i++)
                    {
                        vCellName = drAvgPayment.GetName(i);
                        wsExcel.GetRow(vLinesNo).CreateCell(i - 1);
                        if ((vCellName == "GivCash") ||
                            (vCellName == "LaiAMT") ||
                            (vCellName == "LiFee") ||
                            (vCellName == "HiAMT") ||
                            (vCellName == "HiFee"))
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i - 1).SetCellType(CellType.Numeric);
                            wsExcel.GetRow(vLinesNo).GetCell(i - 1).SetCellValue(double.Parse(drAvgPayment[i].ToString()));
                        }
                        else
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i - 1).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i - 1).SetCellValue(drAvgPayment[i].ToString());
                        }
                        wsExcel.GetRow(vLinesNo).GetCell(i - 1).CellStyle = csData;
                    }
                    vLinesNo++;
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
                    string vRecordNote = "匯出檔案：計算平均薪資" + Environment.NewLine +
                                         "CalAvgPayment.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    //先判斷是不是IE
                    HttpContext.Current.Response.ContentType = "application/octet-stream";
                    HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                    string TourVerision = brObject.Type;
                    string vFileName = "CalAvgPayment";
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
                //throw;
            }
        }
    }
}