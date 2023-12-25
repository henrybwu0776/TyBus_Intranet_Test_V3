using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class ImportTLD_Data : Page
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
        private string vErrorMessage = "";
        private DataTable dtTarget = new DataTable();

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

                UnobtrusiveValidationMode = UnobtrusiveValidationMode.None;

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        plShowData.Visible = false;
                    }
                    else
                    {

                    }
                    string vTempDateURL = "InputDate.aspx?TextboxID=" + eClearDate.ClientID;
                    string vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eClearDate.Attributes["onClick"] = vTempDateScript;
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

        private DataTable CreateTempTable_TLD()
        {
            DataTable dtResult = new DataTable();
            DataColumn dcIndexNo = new DataColumn("IndexNo", typeof(Int32));
            dtResult.Columns.Add(dcIndexNo);
            //檔頭標籤，固定放 "D"
            DataColumn dcTitleMark = new DataColumn("TitleMark", typeof(String));
            dcTitleMark.MaxLength = 1; //資料長度
            dcTitleMark.AllowDBNull = false; //是否可以為空值，False 表示不可以為空值
            dtResult.Columns.Add(dcTitleMark);
            //資料屬性，固定放 "01"
            DataColumn dcDataType = new DataColumn("DataType", typeof(String));
            dcDataType.MaxLength = 2; //資料長度
            dcDataType.AllowDBNull = false; //是否可以為空值，False 表示不可以為空值
            dtResult.Columns.Add(dcDataType);
            //交易模式，01--扣款，04--特種票使用
            DataColumn dcTradeMode = new DataColumn("TradeMode", typeof(String));
            dcTradeMode.MaxLength = 2; //資料長度
            dtResult.Columns.Add(dcTradeMode);
            //清分日期
            DataColumn dcClearDate = new DataColumn("ClearDate", typeof(DateTime));
            dtResult.Columns.Add(dcClearDate);
            //營運日期
            DataColumn dcServiceDate = new DataColumn("ServiceDate", typeof(DateTime));
            dtResult.Columns.Add(dcServiceDate);
            //班別代碼
            DataColumn dcStationCode = new DataColumn("StationCode", typeof(String));
            dcStationCode.MaxLength = 3; //資料長度
            dtResult.Columns.Add(dcStationCode);
            //設備代碼
            DataColumn dcMachineCode = new DataColumn("MachineCode", typeof(String));
            dcMachineCode.MaxLength = 9; //資料長度
            dtResult.Columns.Add(dcMachineCode);
            //車號
            DataColumn dcCar_ID = new DataColumn("Car_ID", typeof(String));
            dcCar_ID.MaxLength = 10; //資料長度
            dtResult.Columns.Add(dcCar_ID);
            //路線代碼
            DataColumn dcLinesNo = new DataColumn("LinesNo", typeof(String));
            dcLinesNo.MaxLength = 5; //資料長度
            dtResult.Columns.Add(dcLinesNo);
            //司機代碼
            DataColumn dcDriverNo = new DataColumn("DriverNo", typeof(String));
            dcDriverNo.MaxLength = 5; //資料長度
            dtResult.Columns.Add(dcDriverNo);
            //開班時間
            DataColumn dcStartTime = new DataColumn("StartTime", typeof(DateTime));
            dtResult.Columns.Add(dcStartTime);
            //身份別
            DataColumn dcIDType = new DataColumn("IDType", typeof(String));
            dcIDType.MaxLength = 2; //資料長度
            dtResult.Columns.Add(dcIDType);
            //晶片卡號
            DataColumn dcChipCardNo = new DataColumn("ChipCardNo", typeof(String));
            dcChipCardNo.MaxLength = 16; //資料長度
            dtResult.Columns.Add(dcChipCardNo);
            //特種票票種
            DataColumn dcSP_Card = new DataColumn("SP_Card", typeof(String));
            dcSP_Card.MaxLength = 3; //資料長度
            dtResult.Columns.Add(dcSP_Card);
            //交易日期時間
            DataColumn dcTradeDate = new DataColumn("TradeDate", typeof(DateTime));
            dtResult.Columns.Add(dcTradeDate);
            //交易序號
            DataColumn dcTradeCode = new DataColumn("TradeCode", typeof(String));
            dcTradeCode.MaxLength = 3; //資料長度
            dtResult.Columns.Add(dcTradeCode);
            //交易金額
            DataColumn dcAmount = new DataColumn("Amount", typeof(Double));
            dtResult.Columns.Add(dcAmount);
            //轉乘優惠金額
            DataColumn dcTransDiscount = new DataColumn("TransDiscount", typeof(Double));
            dtResult.Columns.Add(dcTransDiscount);
            //個人優惠金額
            DataColumn dcPersonalDiscount = new DataColumn("PersonalDiscount", typeof(Double));
            dtResult.Columns.Add(dcPersonalDiscount);
            //出站
            DataColumn dcLeaveStation = new DataColumn("LeaveStation", typeof(String));
            dcLeaveStation.MaxLength = 3; //資料長度
            dtResult.Columns.Add(dcLeaveStation);
            //上、下車旗標
            DataColumn dcIn_OutFlag = new DataColumn("In_OutFlag", typeof(String));
            dcIn_OutFlag.MaxLength = 2; //資料長度
            dtResult.Columns.Add(dcIn_OutFlag);
            //區碼別
            DataColumn dcAreaCode = new DataColumn("AreaCode", typeof(String));
            dcAreaCode.MaxLength = 3; //資料長度
            dtResult.Columns.Add(dcAreaCode);
            //進站
            DataColumn dcInStation = new DataColumn("InStation", typeof(String));
            dcInStation.MaxLength = 3; //資料長度
            dtResult.Columns.Add(dcInStation);
            //2023.12.18 悠遊卡 TLD 檔案新增【錯誤代碼】欄位 
            //錯誤代碼
            DataColumn dcErrorCode = new DataColumn("ErrorCode", typeof(String));
            dcErrorCode.MaxLength = 12; //資料長度
            dtResult.Columns.Add(dcErrorCode);

            return dtResult;
        }

        private DataTable CreateTempTable_ACER()
        {
            DataTable dtResult = new DataTable();
            DataColumn dcIndexNo = new DataColumn("IndexNo", typeof(Int32));
            dtResult.Columns.Add(dcIndexNo);
            DataColumn dcTicketType = new DataColumn("TicketType", typeof(String));//票證公司
            dcTicketType.MaxLength = 12; //資料長度
            dtResult.Columns.Add(dcTicketType);
            DataColumn dcClearDate = new DataColumn("ClearDate", typeof(DateTime));//清分日期
            dtResult.Columns.Add(dcClearDate);
            DataColumn dcServiceDate = new DataColumn("ServiceDate", typeof(DateTime));//營運日期
            dtResult.Columns.Add(dcServiceDate);
            DataColumn dcStationCode = new DataColumn("StationCode", typeof(String));//場站代號
            dcStationCode.MaxLength = 3; //資料長度
            dtResult.Columns.Add(dcStationCode);
            DataColumn dcCar_ID = new DataColumn("Car_ID", typeof(String));//車號
            dcCar_ID.MaxLength = 10; //資料長度
            dtResult.Columns.Add(dcCar_ID);
            DataColumn dcDriverNo = new DataColumn("DriverNo", typeof(String));//駕駛員代號
            dcDriverNo.MaxLength = 10; //資料長度
            dtResult.Columns.Add(dcDriverNo);
            DataColumn dcLinesNo = new DataColumn("LinesNo", typeof(String));//路線
            dcLinesNo.MaxLength = 6; //資料長度
            dtResult.Columns.Add(dcLinesNo);
            DataColumn dcTSAmount = new DataColumn("TSAmount", typeof(Double));//總營運金額	
            dtResult.Columns.Add(dcTSAmount);
            DataColumn dcTNAmount = new DataColumn("TNAmount", typeof(Double));//合計(不含優待學生卡)
            dtResult.Columns.Add(dcTNAmount);
            DataColumn dcStudentAmount = new DataColumn("StudentAmount", typeof(Double));//學生卡金額
            dtResult.Columns.Add(dcStudentAmount);
            DataColumn dcStuAmount_All = new DataColumn("StuAmount_All", typeof(Double));//學生卡金額(全額)
            dtResult.Columns.Add(dcStuAmount_All);
            DataColumn dcStuAmount_Disc = new DataColumn("StuAmount_Disc", typeof(Double));//學生卡金額(優待)
            dtResult.Columns.Add(dcStuAmount_Disc);
            DataColumn dcStuAmount_8KM = new DataColumn("StuAmount_8KM", typeof(Double));//學生卡金額(8KM優惠)
            dtResult.Columns.Add(dcStuAmount_8KM);
            DataColumn dcOldAmount_N = new DataColumn("OldAmount_N", typeof(Double));//敬老愛心卡金額N
            dtResult.Columns.Add(dcOldAmount_N);
            DataColumn dcOldAmount_O = new DataColumn("OldAmount_O", typeof(Double));//敬老愛心卡金額(全額)O
            dtResult.Columns.Add(dcOldAmount_O);
            DataColumn dcOldAmount_P = new DataColumn("OldAmount_P", typeof(Double));//敬老愛心卡金額(優待)P
            dtResult.Columns.Add(dcOldAmount_P);
            DataColumn dcOldAmount_8KM = new DataColumn("OldAmount_8KM", typeof(Double));//敬老愛心卡金額(8KM優惠)
            dtResult.Columns.Add(dcOldAmount_8KM);
            DataColumn dcNormalAmount_8KM = new DataColumn("NormalAmount_8KM", typeof(Double));//普卡金額(8KM優惠)
            dtResult.Columns.Add(dcNormalAmount_8KM);
            DataColumn dcTNAmount_8KM = new DataColumn("TNAmount_8KM", typeof(Double));//合計(8KM優惠)
            dtResult.Columns.Add(dcTNAmount_8KM);
            DataColumn dcTotalAmount = new DataColumn("TotalAmount", typeof(Double));//總合計
            dtResult.Columns.Add(dcTotalAmount);
            DataColumn dcTotalPassCount = new DataColumn("TotalPassCount", typeof(Int32));//總載客人數
            dtResult.Columns.Add(dcTotalPassCount);
            DataColumn dcPassCount_N = new DataColumn("PassCount_N", typeof(Int32));//普卡載客人數
            dtResult.Columns.Add(dcPassCount_N);
            DataColumn dcPassCount_U = new DataColumn("PassCount_U", typeof(Int32));//聯營卡載客人數
            dtResult.Columns.Add(dcPassCount_U);
            DataColumn dcPassCount_S = new DataColumn("PassCount_S", typeof(Int32));//學生卡載客人數
            dtResult.Columns.Add(dcPassCount_S);
            DataColumn dcPassCount_SA = new DataColumn("PassCount_SA", typeof(Int32));//學生卡載客人數(全額)
            dtResult.Columns.Add(dcPassCount_SA);
            DataColumn dcPassCount_SD = new DataColumn("PassCount_SD", typeof(Int32));//學生卡載客人數(優待)
            dtResult.Columns.Add(dcPassCount_SD);
            DataColumn dcPassCount_O = new DataColumn("PassCount_O", typeof(Int32));//敬老愛心卡載客人數AA
            dtResult.Columns.Add(dcPassCount_O);
            DataColumn dcPassCount_OA = new DataColumn("PassCount_OA", typeof(Int32));//敬老愛心卡載客人數(全額)AB
            dtResult.Columns.Add(dcPassCount_OA);
            DataColumn dcPassCount_OD = new DataColumn("PassCount_OD", typeof(Int32));//敬老愛心卡載客人數(優待)AC
            dtResult.Columns.Add(dcPassCount_OD);
            DataColumn dcPassCount_OP = new DataColumn("PassCount_OP", typeof(Int32));//敬老愛心卡扣點筆數
            dtResult.Columns.Add(dcPassCount_OP);
            DataColumn dcPassCount_S8KM = new DataColumn("PassCount_S8KM", typeof(Int32));//8KM優惠學生卡載客人數
            dtResult.Columns.Add(dcPassCount_S8KM);
            DataColumn dcPassCount_O8KM = new DataColumn("PassCount_O8KM", typeof(Int32));//8KM優惠敬老愛心卡載客人數
            dtResult.Columns.Add(dcPassCount_O8KM);
            DataColumn dcPassCount_N8KM = new DataColumn("PassCount_N8KM", typeof(Int32));//8KM優惠普卡載客人數
            dtResult.Columns.Add(dcPassCount_N8KM);
            DataColumn dcPassCount_T8KM = new DataColumn("PassCount_T8KM", typeof(Int32));//8KM優惠總人數
            dtResult.Columns.Add(dcPassCount_T8KM);
            DataColumn dcTransAmount_N = new DataColumn("TransAmount_N", typeof(Double));//轉乘優惠普卡金額
            dtResult.Columns.Add(dcTransAmount_N);
            DataColumn dcTransAmount_S = new DataColumn("TransAmount_S", typeof(Double));//轉乘優惠學生卡金額
            dtResult.Columns.Add(dcTransAmount_S);
            DataColumn dcTransAmount_O = new DataColumn("TransAmount_O", typeof(Double));//轉乘優惠敬老愛心卡金額
            dtResult.Columns.Add(dcTransAmount_O);
            DataColumn dcTransCount_N = new DataColumn("TransCount_N", typeof(Int32));//轉乘優惠普卡載客人數
            dtResult.Columns.Add(dcTransCount_N);
            DataColumn dcTransCount_S = new DataColumn("TransCount_S", typeof(Int32));//轉乘優惠學生卡載客人數
            dtResult.Columns.Add(dcTransCount_S);
            DataColumn dcTransCount_O = new DataColumn("TransCount_O", typeof(Int32));//轉乘優惠敬老愛心卡載客人數
            dtResult.Columns.Add(dcTransCount_O);
            DataColumn dcTicketDiffAmount_N = new DataColumn("TicketDiffAmount_N", typeof(Double));//票差優惠普卡金額
            dtResult.Columns.Add(dcTicketDiffAmount_N);
            DataColumn dcTicketDiffAmount_S = new DataColumn("TicketDiffAmount_S", typeof(Double));//票差優惠學生卡金額
            dtResult.Columns.Add(dcTicketDiffAmount_S);
            DataColumn dcTicketDiffAmount_O = new DataColumn("TicketDiffAmount_O", typeof(Double));//票差優惠敬老愛心卡金額
            dtResult.Columns.Add(dcTicketDiffAmount_O);
            DataColumn dcTicketDiffCount_N = new DataColumn("TicketDiffCount_N", typeof(Double));//票差優惠普卡載客人數
            dtResult.Columns.Add(dcTicketDiffCount_N);
            DataColumn dcTicketDiffCount_S = new DataColumn("TicketDiffCount_S", typeof(Double));//票差優惠學生卡載客人數
            dtResult.Columns.Add(dcTicketDiffCount_S);
            DataColumn dcTicketDiffCount_O = new DataColumn("TicketDiffCount_O", typeof(Double));//票差優惠敬老愛心卡載客人數
            dtResult.Columns.Add(dcTicketDiffCount_O);
            DataColumn dcAreaDisAmount_N = new DataColumn("AreaDisAmount_N", typeof(Double));//區間優惠普卡金額
            dtResult.Columns.Add(dcAreaDisAmount_N);
            DataColumn dcAreaDisAmount_S = new DataColumn("AreaDisAmount_S", typeof(Double));//區間優惠學生卡金額
            dtResult.Columns.Add(dcAreaDisAmount_S);
            DataColumn dcAreaDisAmount_O = new DataColumn("AreaDisAmount_O", typeof(Double));//區間優惠敬老愛心卡金額
            dtResult.Columns.Add(dcAreaDisAmount_O);
            DataColumn dcAreaDisCount_N = new DataColumn("AreaDisCount_N", typeof(Int32));//區間優惠普卡載客人數
            dtResult.Columns.Add(dcAreaDisCount_N);
            DataColumn dcAreaDisCount_S = new DataColumn("AreaDisCount_S", typeof(Int32));//區間優惠學生卡載客人數
            dtResult.Columns.Add(dcAreaDisCount_S);
            DataColumn dcAreaDisCount_O = new DataColumn("AreaDisCount_O", typeof(Int32));//區間優惠敬老愛心卡載客人數
            dtResult.Columns.Add(dcAreaDisCount_O);
            DataColumn dc1280MTicketAmount_N = new DataColumn("1280MTicketAmount_N", typeof(Double));//雙北定期票普卡金額
            dtResult.Columns.Add(dc1280MTicketAmount_N);
            DataColumn dc1280MTicketAmount_S = new DataColumn("1280MTicketAmount_S", typeof(Double));//雙北定期票學生卡金額
            dtResult.Columns.Add(dc1280MTicketAmount_S);
            DataColumn dc1280MTicketAmount_O = new DataColumn("1280MTicketAmount_O", typeof(Double));//雙北定期票敬老愛心卡金額
            dtResult.Columns.Add(dc1280MTicketAmount_O);
            DataColumn dc1280MTicketCount_N = new DataColumn("1280MTicketCount_N", typeof(Double));//雙北定期票普卡載客人數
            dtResult.Columns.Add(dc1280MTicketCount_N);
            DataColumn dc1280MTicketCount_S = new DataColumn("1280MTicketCount_S", typeof(Double));//雙北定期票學生卡載客人數
            dtResult.Columns.Add(dc1280MTicketCount_S);
            DataColumn dc1280MTicketCount_O = new DataColumn("1280MTicketCount_O", typeof(Double));//雙北定期票敬老愛心卡載客人數
            dtResult.Columns.Add(dc1280MTicketCount_O);
            DataColumn dcQRCount = new DataColumn("QRCount", typeof(Int32));//QR人數
            dtResult.Columns.Add(dcQRCount);
            DataColumn dcQRAmount = new DataColumn("QRAmount", typeof(Double));//QR金額
            dtResult.Columns.Add(dcQRAmount);
            DataColumn dcTPassCount = new DataColumn("TPassCount", typeof(Int32));//基北北桃月票人數
            dtResult.Columns.Add(dcTPassCount);
            DataColumn dcTPassAmount = new DataColumn("TPassAmount", typeof(Double));//基北北桃月票金額
            dtResult.Columns.Add(dcTPassAmount);

            return dtResult;
        }

        private DataTable GetDataTableData_TLD(string fTempStr)
        {
            DataTable dtResult = new DataTable();
            dtResult = CreateTempTable_TLD();
            string[] vaTempStr = fTempStr.Split(Environment.NewLine.ToCharArray());
            foreach (string vDataStr in vaTempStr)
            {
                if ((vDataStr.Length > 0) && (vDataStr.Substring(0, 3) == "D01"))
                {
                    DataRow drTemp = dtResult.NewRow();
                    drTemp["TitleMark"] = vDataStr.Substring(0, 1).Trim();
                    drTemp["DataType"] = vDataStr.Substring(1, 2).Trim();
                    drTemp["TradeMode"] = vDataStr.Substring(3, 2).Trim();
                    drTemp["ClearDate"] = DateTime.Parse(PF.TransNoneSymbolDateStrToDate(vDataStr.Substring(5, 8).Trim(), "B") + " 00:00:00");
                    drTemp["ServiceDate"] = DateTime.Parse(PF.TransNoneSymbolDateStrToDate(vDataStr.Substring(13, 8).Trim(), "B") + " 00:00:00");
                    drTemp["StationCode"] = vDataStr.Substring(21, 3).Trim();
                    drTemp["MachineCode"] = vDataStr.Substring(24, 9).Trim();
                    drTemp["Car_ID"] = vDataStr.Substring(33, 10).Trim().Replace("_B", "");
                    drTemp["LinesNo"] = vDataStr.Substring(43, 5).Trim();
                    drTemp["DriverNo"] = vDataStr.Substring(48, 5).Trim();
                    drTemp["StartTime"] = DateTime.Parse(PF.TransNoneSymbolDateStrToDate(vDataStr.Substring(53, 8).Trim(), "B") + " " + vDataStr.Substring(61, 2).Trim() + ":" + vDataStr.Substring(63, 2).Trim() + ":" + vDataStr.Substring(65, 2).Trim());
                    drTemp["IDType"] = vDataStr.Substring(67, 2).Trim();
                    drTemp["ChipCardNo"] = vDataStr.Substring(69, 16).Trim();
                    drTemp["SP_Card"] = vDataStr.Substring(85, 3).Trim();
                    drTemp["TradeDate"] = DateTime.Parse(PF.TransNoneSymbolDateStrToDate(vDataStr.Substring(88, 8).Trim(), "B") + " " + vDataStr.Substring(96, 2).Trim() + ":" + vDataStr.Substring(98, 2).Trim() + ":" + vDataStr.Substring(100, 2).Trim());
                    drTemp["TradeCode"] = vDataStr.Substring(102, 3).Trim();
                    drTemp["Amount"] = Double.Parse(vDataStr.Substring(105, 5).Trim());
                    drTemp["TransDiscount"] = Double.Parse(vDataStr.Substring(110, 5).Trim());
                    drTemp["PersonalDiscount"] = Double.Parse(vDataStr.Substring(115, 5).Trim());
                    drTemp["LeaveStation"] = vDataStr.Substring(120, 3).Trim();
                    drTemp["In_OutFlag"] = vDataStr.Substring(123, 2).Trim();
                    drTemp["AreaCode"] = vDataStr.Substring(125, 3).Trim();
                    drTemp["InStation"] = vDataStr.Substring(128, 3).Trim();
                    //2023.12.18 悠遊卡 TLD 檔案新增【錯誤代碼】欄位 
                    drTemp["ErrorCode"] = String.Empty; //因為是舊版沒有帶錯誤代碼的資料，所以這個欄位先帶空值
                    dtResult.Rows.Add(drTemp);
                }
            }
            return dtResult;
        }

        private DataTable GetDataTableData_TLD_E(string fTempStr)
        {
            DataTable dtResult = new DataTable();
            dtResult = CreateTempTable_TLD();
            string[] vaTempStr = fTempStr.Split(Environment.NewLine.ToCharArray());
            string vErrorCode = "";
            string vErrorCode_Temp = "";
            int vErrorSourceInt = 0;
            string vTempStr = "";
            foreach (string vDataStr in vaTempStr)
            {
                if ((vDataStr.Length > 0) && (vDataStr.Substring(0, 3) == "D01"))
                {
                    try
                    {
                        DataRow drTemp = dtResult.NewRow();
                        drTemp["TitleMark"] = vDataStr.Substring(0, 1).Trim();
                        drTemp["DataType"] = vDataStr.Substring(1, 2).Trim();
                        drTemp["TradeMode"] = vDataStr.Substring(3, 2).Trim();
                        drTemp["ClearDate"] = DateTime.Parse(PF.TransNoneSymbolDateStrToDate(vDataStr.Substring(5, 8).Trim(), "B") + " 00:00:00");
                        drTemp["ServiceDate"] = DateTime.Parse(PF.TransNoneSymbolDateStrToDate(vDataStr.Substring(13, 8).Trim(), "B") + " 00:00:00");
                        drTemp["StationCode"] = vDataStr.Substring(21, 3).Trim();
                        drTemp["MachineCode"] = vDataStr.Substring(24, 9).Trim();
                        drTemp["Car_ID"] = vDataStr.Substring(33, 10).Trim().Replace("_B", "");
                        drTemp["LinesNo"] = vDataStr.Substring(43, 5).Trim();
                        drTemp["DriverNo"] = vDataStr.Substring(48, 5).Trim();
                        drTemp["StartTime"] = DateTime.Parse(PF.TransNoneSymbolDateStrToDate(vDataStr.Substring(53, 8).Trim(), "B") + " " + vDataStr.Substring(61, 2).Trim() + ":" + vDataStr.Substring(63, 2).Trim() + ":" + vDataStr.Substring(65, 2).Trim());
                        drTemp["IDType"] = vDataStr.Substring(67, 2).Trim();
                        drTemp["ChipCardNo"] = vDataStr.Substring(69, 16).Trim();
                        drTemp["SP_Card"] = vDataStr.Substring(85, 3).Trim();
                        drTemp["TradeDate"] = DateTime.Parse(PF.TransNoneSymbolDateStrToDate(vDataStr.Substring(88, 8).Trim(), "B") + " " + vDataStr.Substring(96, 2).Trim() + ":" + vDataStr.Substring(98, 2).Trim() + ":" + vDataStr.Substring(100, 2).Trim());
                        drTemp["TradeCode"] = vDataStr.Substring(102, 3).Trim();
                        drTemp["Amount"] = Double.Parse(vDataStr.Substring(105, 5).Trim());
                        drTemp["TransDiscount"] = Double.Parse(vDataStr.Substring(110, 5).Trim());
                        drTemp["PersonalDiscount"] = Double.Parse(vDataStr.Substring(115, 5).Trim());
                        drTemp["LeaveStation"] = vDataStr.Substring(120, 3).Trim();
                        drTemp["In_OutFlag"] = vDataStr.Substring(123, 2).Trim();
                        drTemp["AreaCode"] = vDataStr.Substring(125, 3).Trim();
                        drTemp["InStation"] = vDataStr.Substring(128, 3).Trim();
                        //2023.12.18 悠遊卡 TLD 檔案新增【錯誤代碼】欄位 
                        vErrorCode_Temp = vDataStr.Substring(131, 6).Trim();
                        if (vErrorCode_Temp.Trim().Substring(0, 1).ToUpper() == "A") //A類錯誤代碼要另外處理
                        {
                            if (Int32.TryParse(vErrorCode_Temp.Trim().Substring(1), out vErrorSourceInt)) //試著去掉開頭的 "A" 字，轉成整數數值
                            {
                                vTempStr = ("00000" + Convert.ToString(vErrorSourceInt, 2));
                                vErrorCode = "A" + vTempStr.Substring(vTempStr.Length - 5, 5);
                            }
                        }
                        drTemp["ErrorCode"] = (vErrorCode_Temp.Trim() == "000000") ? "" :
                                              (vErrorCode_Temp.Trim().Substring(0, 1).ToUpper() == "B") ? vErrorCode_Temp.Trim() : vErrorCode.Trim();
                        dtResult.Rows.Add(drTemp);
                    }
                    catch (Exception eMessage)
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('" + eMessage.Message + "')");
                        Response.Write("</" + "Script>");
                    }

                }
            }
            return dtResult;
        }

        private DataTable GetDataTableData_ACER()
        {
            DataTable dtResult = new DataTable();
            string vExtName = Path.GetExtension(fuFileName.FileName);

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (vExtName)
            {
                case ".xls":
                    dtResult = GetExcelData_H();
                    break;

                case ".xlsx":
                    dtResult = GetExcelData_X();
                    break;
            }
            return dtResult;
        }

        private DataTable GetExcelData_H()
        {
            DataTable dtResult = new DataTable();
            HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuFileName.FileContent);
            HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
            if (sheetExcel_H.LastRowNum > 0)
            {
                dtResult = CreateTempTable_ACER();
                for (int i = sheetExcel_H.FirstRowNum + 1; i <= sheetExcel_H.LastRowNum; i++)
                {
                    HSSFRow rowExcel_H = (HSSFRow)sheetExcel_H.GetRow(i);
                    if ((rowExcel_H != null) && (rowExcel_H.GetCell(0).StringCellValue.Trim() != "票證公司"))
                    {
                        DataRow rowSQL = dtResult.NewRow();
                        for (int j = 0; j < rowExcel_H.Cells.Count; j++)
                        {
                            rowSQL[j + 1] = rowExcel_H.GetCell(j).ToString().Trim();
                        }
                        dtResult.Rows.Add(rowSQL);
                    }
                }
            }
            return dtResult;
        }

        private DataTable GetExcelData_X()
        {
            DataTable dtResult = new DataTable();
            XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuFileName.FileContent);
            XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
            if (sheetExcel_X.LastRowNum > 0)
            {
                dtResult = CreateTempTable_ACER();
                for (int i = sheetExcel_X.FirstRowNum + 1; i <= sheetExcel_X.LastRowNum; i++)
                {
                    XSSFRow rowExcel_X = (XSSFRow)sheetExcel_X.GetRow(i);
                    if ((rowExcel_X != null) && (rowExcel_X.GetCell(0).StringCellValue.Trim() != "票證公司"))
                    {
                        DataRow rowSQL = dtResult.NewRow();
                        for (int j = 0; j < rowExcel_X.Cells.Count; j++)
                        {
                            rowSQL[j + 1] = rowExcel_X.GetCell(j).ToString().Trim();
                        }
                        dtResult.Rows.Add(rowSQL);
                    }
                }
            }
            return dtResult;
        }

        /* 預覽改查詢
        protected void bbPreView_Click(object sender, EventArgs e)
        {
            if (fuFileName.FileName.Trim() != "")
            {
                StreamReader srTemp = new StreamReader(fuFileName.PostedFile.InputStream);
                string vTempStr = srTemp.ReadToEnd();
                if (vTempStr.Length > 0)
                {
                    dtTarget = GetDataTableData(vTempStr);
                    gvTLD_Data.DataSourceID = "";
                    gvTLD_Data.DataSource = dtTarget;
                    plShowData.Visible = true;
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇檔案！')");
                Response.Write("</" + "Script>");
            }
        } //*/

        protected void bbImport_Click(object sender, EventArgs e)
        {
            string vTableName = "";
            if (fuFileName.FileName.Trim() != "")
            {
                string vTempStr = "";
                switch (eImportSourceType.SelectedValue.Trim())
                {
                    case "TLD_Data":
                        StreamReader srTemp = new StreamReader(fuFileName.PostedFile.InputStream);
                        vTempStr = srTemp.ReadToEnd();
                        vTableName = "TLD_Data";
                        dtTarget = GetDataTableData_TLD(vTempStr);
                        break;

                    case "TLD_Data2":
                        //2023.12.18 悠遊卡 TLD 檔案新增【錯誤代碼】欄位 
                        StreamReader srTemp_E = new StreamReader(fuFileName.PostedFile.InputStream);
                        vTempStr = srTemp_E.ReadToEnd();
                        vTableName = "TLD_Data";
                        dtTarget = GetDataTableData_TLD_E(vTempStr);
                        break;

                    case "ACER_Data":
                        vTableName = "ACER_Data";
                        dtTarget = GetDataTableData_ACER();
                        break;
                }
                if ((dtTarget != null) && (dtTarget.Rows.Count > 0))
                {
                    using (SqlBulkCopy sbkTemp = new SqlBulkCopy(vConnStr, SqlBulkCopyOptions.UseInternalTransaction))
                    {
                        try
                        {
                            //sbkTemp.DestinationTableName = "TLD_Data";
                            sbkTemp.DestinationTableName = vTableName;
                            sbkTemp.WriteToServer(dtTarget);
                            dtTarget.Clear();
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('資料匯入完畢!')");
                            Response.Write("</" + "Script>");
                        }
                        catch (Exception eMessage)
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('" + eMessage.Message + "')");
                            Response.Write("</" + "Script>");
                        }
                    }

                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('無可用資料！')");
                    Response.Write("</" + "Script>");
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇檔案！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 畫面查詢資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSearch_Click(object sender, EventArgs e)
        {
            plShowData.Visible = false;
            gvTLD_Data.Visible = false;
            gvACER_Data.Visible = false;
            DateTime vClearDate_Search;
            if (DateTime.TryParse(eClearDate.Text.Trim(), out vClearDate_Search))
            {
                switch (eImportSourceType.SelectedValue.Trim())
                {
                    case "TLD_Data":
                    case "TLD_Data2":
                        dsTLD_Data.SelectCommand = "SELECT TitleMark, DataType, TradeMode, ClearDate, ServiceDate, StationCode, MachineCode, " + Environment.NewLine +
                                                   "       Car_ID, LinesNo, DriverNo, StartTime, IDType, ChipCardNo, SP_Card, TradeDate, TradeCode, " + Environment.NewLine +
                                                   "       Amount, TransDiscount, PersonalDiscount, LeaveStation, In_OutFlag, AreaCode, InStation, ErrorCode " + Environment.NewLine +
                                                   "  FROM TLD_Data " + Environment.NewLine +
                                                   " WHERE ClearDate = @ClearDate ";
                        dsTLD_Data.SelectParameters.Clear();
                        dsTLD_Data.SelectParameters.Add(new Parameter("ClearDate", DbType.Date, vClearDate_Search.ToString("yyyy/MM/dd")));
                        dsTLD_Data.DataBind();
                        gvTLD_Data.DataBind();
                        plShowData.Visible = true;
                        gvTLD_Data.Visible = true;
                        break;

                    case "ACER_Data":
                        dsACER_Data.SelectCommand = "SELECT IndexNo, TicketType, ClearDate, ServiceDate, StationCode, Car_ID, DriverNo, LinesNo, " + Environment.NewLine +
                                                    "       TSAmount, StudentAmount, OldAmount_N, NormalAmount_8KM, TotalAmount, TotalPassCount, " + Environment.NewLine +
                                                    "       PassCount_N, PassCount_U, PassCount_S, PassCount_O, TPassCount, TPassAmount " + Environment.NewLine +
                                                    "  FROM ACER_Data " + Environment.NewLine +
                                                    " WHERE ClearDate = @ClearDate " + Environment.NewLine +
                                                    " ORDER BY IndexNo ";
                        dsACER_Data.SelectParameters.Clear();
                        dsACER_Data.SelectParameters.Add(new Parameter("ClearDate", DbType.Date, vClearDate_Search.ToString("yyyy/MM/dd")));
                        dsACER_Data.DataBind();
                        gvACER_Data.DataBind();
                        plShowData.Visible = true;
                        gvACER_Data.Visible = true;
                        break;
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請輸入清分日期！')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 匯出資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExport_Click(object sender, EventArgs e)
        {
            plShowData.Visible = false;
            gvTLD_Data.Visible = false;
            gvACER_Data.Visible = false;
            DateTime vClearDate_Export;
            string vSelectStr = "";
            DataTable dtTemp = new DataTable();

            if (DateTime.TryParse(eClearDate.Text.Trim(), out vClearDate_Export))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                switch (eImportSourceType.SelectedValue.Trim())
                {
                    case "TLD_Data":
                    case "TLD_Data2":
                        vSelectStr = "SELECT TitleMark, DataType, TradeMode, ClearDate, ServiceDate, StationCode, MachineCode, Car_ID, LinesNo, DriverNo, StartTime, IDType, " + Environment.NewLine +
                                     "       ChipCardNo, SP_Card, TradeDate, TradeCode, Amount, TransDiscount, PersonalDiscount, LeaveStation, In_OutFlag, AreaCode, InStation, ErrorCode " + Environment.NewLine +
                                     "  FROM TLD_Data " + Environment.NewLine +
                                      " WHERE ClearDate = @ClearDate";
                        break;

                    case "ACER_Data":
                        vSelectStr = "SELECT TicketType, ClearDate, ServiceDate, StationCode, Car_ID, DriverNo, LinesNo, TSAmount, TNAmount, StudentAmount, StuAmount_All, " + Environment.NewLine +
                                     "       StuAmount_Disc, StuAmount_8KM, OldAmount_N, OldAmount_O, OldAmount_P, OldAmount_8KM, NormalAmount_8KM, TNAmount_8KM, TotalAmount, " + Environment.NewLine +
                                     "       TotalPassCount, PassCount_N, PassCount_U, PassCount_S, PassCount_SA, PassCount_SD, PassCount_O, PassCount_OA, PassCount_OD, " + Environment.NewLine +
                                     "       PassCount_OP, PassCount_S8KM, PassCount_O8KM, PassCount_N8KM, PassCount_T8KM, TransAmount_N, TransAmount_S, TransAmount_O, " + Environment.NewLine +
                                     "       TransCount_N, TransCount_S, TransCount_O, TicketDiffAmount_N, TicketDiffAmount_S, TicketDiffAmount_O, TicketDiffCount_N, " + Environment.NewLine +
                                     "       TicketDiffCount_S, TicketDiffCount_O, AreaDisAmount_N, AreaDisAmount_S, AreaDisAmount_O, AreaDisCount_N, AreaDisCount_S, " + Environment.NewLine +
                                     "       AreaDisCount_O, 1280MTicketAmount_N, 1280MTicketAmount_S, 1280MTicketAmount_O, 1280MTicketCount_N, 1280MTicketCount_S, " + Environment.NewLine +
                                     "       1280MTicketCount_O, QRCount, QRAmount, TPassCount, TPassAmount " + Environment.NewLine +
                                     "  FROM ACER_Data " + Environment.NewLine +
                                     " WHERE ClearDate = @ClearDate";
                        break;
                }
                using (SqlConnection connExport = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdExport = new SqlCommand(vSelectStr, connExport);
                    SqlParameter parTemp = new SqlParameter("ClearDate", SqlDbType.DateTime);
                    parTemp.Value = vClearDate_Export.ToString("yyyy/MM/dd");
                    cmdExport.Parameters.Clear();
                    cmdExport.Parameters.Add(parTemp);
                    connExport.Open();
                    SqlDataReader drExport = cmdExport.ExecuteReader();
                    if (drExport.HasRows)
                    {
                        ExportFile(drExport, vClearDate_Export);
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('指定日期無清分資料！')");
                        Response.Write("</" + "Script>");
                    }
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請輸入清分日期！')");
                Response.Write("</" + "Script>");
            }
        }

        private string GetErrorCodeA(string fErrorCode)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vResultStr = "";
            string vFKey = "匯入交易明細資料TLD_Data        ErrorCodeA";
            int vTempInt = 0;
            for (int i = 1; i < 6; i++)
            {
                if ((fErrorCode.Substring(i, 1) == "1") && (vResultStr.Trim() == ""))
                {
                    vTempInt = (int)Math.Pow(2, 5 - i);
                    vResultStr = PF.GetValue(vConnStr, "select ClassTxt from DBDICB where ClassNo = '" + vTempInt.ToString() + "' and FKey = '" + vFKey + "' ", "ClassTxt");
                }
                else if ((fErrorCode.Substring(i, 1) == "1") && (vResultStr.Trim() != ""))
                {
                    vTempInt = (int)Math.Pow(2, 5 - i);
                    vResultStr = vResultStr + ", " + PF.GetValue(vConnStr, "select ClassTxt from DBDICB where ClassNo = '" + vTempInt.ToString() + "' and FKey = '" + vFKey + "' ", "ClassTxt");
                }
            }
            return vResultStr;
        }

        private string GetErrorCodeB(string fErrorCode)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vResultStr = "";
            string vFKey = "匯入交易明細資料TLD_Data        ErrorCodeB";
            string vClassNo = Int32.Parse(fErrorCode.Substring(1, 5)).ToString();
            vResultStr = PF.GetValue(vConnStr, "select ClassTxt from DBDICB where ClassNo = '" + vClassNo + "' and FKey = '" + vFKey + "' ", "ClassTxt");
            return vResultStr;
        }

        private void ExportFile(SqlDataReader drTemp, DateTime fClearDate)
        {
            XSSFWorkbook wbExcel = new XSSFWorkbook();
            //Excel 工作表
            XSSFSheet wsExcel;
            //設定標題欄位的格式
            XSSFCellStyle csTitle = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中

            //設定字體格式
            XSSFFont fontTitle = (XSSFFont)wbExcel.CreateFont();
            //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
            fontTitle.IsBold = true;
            fontTitle.FontHeightInPoints = 12; //字體大小
            csTitle.SetFont(fontTitle);

            //設定資料內容欄位的格式
            XSSFCellStyle csData = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csData.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            XSSFFont fontData = (XSSFFont)wbExcel.CreateFont();
            fontData.FontHeightInPoints = 12;
            csData.SetFont(fontData);

            XSSFCellStyle csData_Int = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            XSSFDataFormat format = (XSSFDataFormat)wbExcel.CreateDataFormat();
            csData_Int.DataFormat = format.GetFormat("###,##0");

            XSSFCellStyle csData_Float = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Float.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csData_Float.DataFormat = format.GetFormat("###,##0.00");

            string vHeaderText = "";
            int vLinesNo = 0;
            string vFileName = fClearDate.Year.ToString("D4") + "年" +
                               fClearDate.Month.ToString("D2") + "月" +
                               fClearDate.Day.ToString("D2") + "日" +
                               eImportSourceType.SelectedItem.Text.Trim();
            DateTime vDate;
            string vFieldValue = "";
            string vErrorCode = "";
            double vTempFloat = 0.0;

            //新增一個工作表
            wsExcel = (XSSFSheet)wbExcel.CreateSheet(vFileName);
            //寫入標題列
            vLinesNo = 0;
            wsExcel.CreateRow(vLinesNo);
            for (int i = 0; i < drTemp.FieldCount; i++)
            {
                vHeaderText = (drTemp.GetName(i).ToUpper().Trim() == "TITLEMARK") ? "檔頭標籤" :
                              (drTemp.GetName(i).ToUpper().Trim() == "DATATYPE") ? "資料屬性" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TRADEMODE") ? "交易模式" :
                              (drTemp.GetName(i).ToUpper().Trim() == "CLEARDATE") ? "清分日期" :
                              (drTemp.GetName(i).ToUpper().Trim() == "SERVICEDATE") ? "營運日期" :
                              (drTemp.GetName(i).ToUpper().Trim() == "STATIONCODE") ? "站別代碼" :
                              (drTemp.GetName(i).ToUpper().Trim() == "MACHINECODE") ? "設備代碼" :
                              (drTemp.GetName(i).ToUpper().Trim() == "CAR_ID") ? "車號" :
                              (drTemp.GetName(i).ToUpper().Trim() == "LINESNO") ? "路線代碼" :
                              (drTemp.GetName(i).ToUpper().Trim() == "DRIVERNO") ? "司機代碼" :
                              (drTemp.GetName(i).ToUpper().Trim() == "STARTTIME") ? "開班時間" :
                              (drTemp.GetName(i).ToUpper().Trim() == "IDTYPE") ? "身份別" :
                              (drTemp.GetName(i).ToUpper().Trim() == "CHIPCARDNO") ? "晶片卡號" :
                              (drTemp.GetName(i).ToUpper().Trim() == "SP_CARD") ? "特種票種" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TRADEDATE") ? "交易日期" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TRADECODE") ? "交易序號" :
                              (drTemp.GetName(i).ToUpper().Trim() == "AMOUNT") ? "交易金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TRANSDISCOUNT") ? "轉乘優惠" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PERSONALDISCOUNT") ? "個人優惠" :
                              (drTemp.GetName(i).ToUpper().Trim() == "LEAVESTATION") ? "出站" :
                              (drTemp.GetName(i).ToUpper().Trim() == "IN_OUTFLAG") ? "上下車" :
                              (drTemp.GetName(i).ToUpper().Trim() == "AREACODE") ? "區碼" :
                              (drTemp.GetName(i).ToUpper().Trim() == "INSTATION") ? "進站" :
                              (drTemp.GetName(i).ToUpper().Trim() == "INDEXNO") ? "序號" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TICKETTYPE") ? "票證公司" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TSAMOUNT") ? "總營運金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "STUDENTAMOUNT") ? "學生卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "OLDAMOUNT_N") ? "敬老愛心卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "NORMALAMOUNT_8KM") ? "普卡金額 (8KM)" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TOTALAMOUNT") ? "總合計" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TOTALPASSCOUNT") ? "總載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PASSCOUNT_N") ? "普卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PASSCOUNT_U") ? "聯營卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PASSCOUNT_S") ? "學生卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PASSCOUNT_O") ? "敬老愛心卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TPASSCOUNT") ? "基北北桃月票人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TPASSAMOUNT") ? "基北北桃月票金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TNAMOUNT") ? "合計(不含優待學生卡)" :
                              (drTemp.GetName(i).ToUpper().Trim() == "STUAMOUNT_ALL") ? "學生卡金額(全額)" :
                              (drTemp.GetName(i).ToUpper().Trim() == "STUAMOUNT_DISC") ? "學生卡金額(優待)" :
                              (drTemp.GetName(i).ToUpper().Trim() == "STUAMOUNT_8KM") ? "學生卡金額(8KM優惠)" :
                              (drTemp.GetName(i).ToUpper().Trim() == "OLDAMOUNT_O") ? "敬老愛心卡金額(全額)" :
                              (drTemp.GetName(i).ToUpper().Trim() == "OLDAMOUNT_P") ? "敬老愛心卡金額(優待)" :
                              (drTemp.GetName(i).ToUpper().Trim() == "OLDAMOUNT_8KM") ? "敬老愛心卡金額(8KM優惠)" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TNAMOUNT_8KM") ? "合計(8KM優惠)" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PASSCOUNT_SA") ? "學生卡載客人數(全額)" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PASSCOUNT_SD") ? "學生卡載客人數(優待)" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PASSCOUNT_OA") ? "敬老愛心卡載客人數(全額)" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PASSCOUNT_OD") ? "敬老愛心卡載客人數(優待)" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PASSCOUNT_OP") ? "敬老愛心卡扣點筆數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PASSCOUNT_S8KM") ? "8KM優惠學生卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PASSCOUNT_O8KM") ? "8KM優惠敬老愛心卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PASSCOUNT_N8KM") ? "8KM優惠普卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "PASSCOUNT_T8KM") ? "8KM優惠總人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TRANSAMOUNT_N") ? "轉乘優惠普卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TRANSAMOUNT_S") ? "轉乘優惠學生卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TRANSAMOUNT_O") ? "轉乘優惠敬老愛心卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TRANSCOUNT_N") ? "轉乘優惠普卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TRANSCOUNT_S") ? "轉乘優惠學生卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TRANSCOUNT_O") ? "轉乘優惠敬老愛心卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TICKETDIFFAMOUNT_N") ? "票差優惠普卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TICKETDIFFAMOUNT_S") ? "票差優惠學生卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TICKETDIFFAMOUNT_O") ? "票差優惠敬老愛心卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TICKETDIFFCOUNT_N") ? "票差優惠普卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TICKETDIFFCOUNT_S") ? "票差優惠學生卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "TICKETDIFFCOUNT_O") ? "票差優惠敬老愛心卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "AREADISAMOUNT_N") ? "區間優惠普卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "AREADISAMOUNT_S") ? "區間優惠學生卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "AREADISAMOUNT_O") ? "區間優惠敬老愛心卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "AREADISCOUNT_N") ? "區間優惠普卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "AREADISCOUNT_S") ? "區間優惠學生卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "AREADISCOUNT_O") ? "區間優惠敬老愛心卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "1280MTICKETAMOUNT_N") ? "雙北定期票普卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "1280MTICKETAMOUNT_S") ? "雙北定期票學生卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "1280MTICKETAMOUNT_O") ? "雙北定期票敬老愛心卡金額" :
                              (drTemp.GetName(i).ToUpper().Trim() == "1280MTICKETCOUNT_N") ? "雙北定期票普卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "1280MTICKETCOUNT_S") ? "雙北定期票學生卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "1280MTICKETCOUNT_O") ? "雙北定期票敬老愛心卡載客人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "QRCOUNT") ? "QR人數" :
                              (drTemp.GetName(i).ToUpper().Trim() == "QRAMOUNT") ? "QR金額" :
                              //2023.12.18 悠遊卡 TLD 檔案新增【錯誤代碼】欄位 
                              (drTemp.GetName(i).ToUpper().Trim() == "ERRORCODE") ? "錯誤代碼" : drTemp.GetName(i).Trim();
                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
            }
            while (drTemp.Read())
            {
                vLinesNo++;
                wsExcel.CreateRow(vLinesNo);
                for (int i = 0; i < drTemp.FieldCount; i++)
                {
                    vHeaderText = drTemp.GetName(i).ToUpper();
                    wsExcel.GetRow(vLinesNo).CreateCell(i);
                    if ((vHeaderText == "TITLEMARK") ||
                        (vHeaderText == "DATATYPE") ||
                        (vHeaderText == "TRADEMODE") ||
                        (vHeaderText == "STATIONCODE") ||
                        (vHeaderText == "MACHINECODE") ||
                        (vHeaderText == "CAR_ID") ||
                        (vHeaderText == "LINESNO") ||
                        (vHeaderText == "IDTYPE") ||
                        (vHeaderText == "CHIPCARDNO") ||
                        (vHeaderText == "SP_CARD") ||
                        (vHeaderText == "TRADECODE") ||
                        (vHeaderText == "LEAVESTATION") ||
                        (vHeaderText == "IN_OUTFLAG") ||
                        (vHeaderText == "AREACODE") ||
                        (vHeaderText == "DRIVERNO") ||
                        (vHeaderText == "INSTATION") ||
                        (vHeaderText == "TICKETTYPE"))
                    {
                        //文字欄位
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drTemp[i].ToString());
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                    }
                    else if (vHeaderText == "ERRORCODE")
                    {
                        //2023.12.18 悠遊卡 TLD 檔案新增【錯誤代碼】欄位 
                        vErrorCode = (vFieldValue.Substring(0, 1).ToUpper() == "A") ? GetErrorCodeA(vFieldValue) :
                                     (vFieldValue.Substring(0, 1).ToUpper() == "B") ? GetErrorCodeB(vFieldValue) : "";
                        vFieldValue = drTemp[i].ToString().Trim();
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vErrorCode);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                    }
                    else if ((vHeaderText == "CLEARDATE") || (vHeaderText == "SERVICEDATE"))
                    {
                        string vTempStr = drTemp[i].ToString();
                        vDate = DateTime.Parse(drTemp[i].ToString());
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vDate.ToString("yyyy/MM/dd"));
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                    }
                    else if ((vHeaderText == "STARTTIME") || (vHeaderText == "TRADEDATE"))
                    {
                        string vTempStr = drTemp[i].ToString();
                        vDate = DateTime.Parse(drTemp[i].ToString());
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vDate.ToString("yyyy/MM/dd HH:mm:ss"));
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                    }
                    else if ((vHeaderText == "INDEXNO") ||
                             (vHeaderText == "TOTALPASSCOUNT") ||
                             (vHeaderText == "PASSCOUNT_N") ||
                             (vHeaderText == "PASSCOUNT_U") ||
                             (vHeaderText == "PASSCOUNT_S") ||
                             (vHeaderText == "PASSCOUNT_O") ||
                             (vHeaderText == "TPASSCOUNT") ||
                             (vHeaderText == "PASSCOUNT_SA") ||
                             (vHeaderText == "PASSCOUNT_SD") ||
                             (vHeaderText == "PASSCOUNT_OA") ||
                             (vHeaderText == "PASSCOUNT_OD") ||
                             (vHeaderText == "PASSCOUNT_OP") ||
                             (vHeaderText == "PASSCOUNT_S8KM") ||
                             (vHeaderText == "PASSCOUNT_O8KM") ||
                             (vHeaderText == "PASSCOUNT_N8KM") ||
                             (vHeaderText == "PASSCOUNT_T8KM") ||
                             (vHeaderText == "TRANSCOUNT_N") ||
                             (vHeaderText == "TRANSCOUNT_S") ||
                             (vHeaderText == "TRANSCOUNT_O") ||
                             (vHeaderText == "TICKETDIFFCOUNT_N") ||
                             (vHeaderText == "TICKETDIFFCOUNT_S") ||
                             (vHeaderText == "TICKETDIFFCOUNT_O") ||
                             (vHeaderText == "AREADISCOUNT_N") ||
                             (vHeaderText == "AREADISCOUNT_S") ||
                             (vHeaderText == "AREADISCOUNT_O") ||
                             (vHeaderText == "1280MTICKETCOUNT_N") ||
                             (vHeaderText == "1280MTICKETCOUNT_S") ||
                             (vHeaderText == "1280MTICKETCOUNT_O") ||
                             (vHeaderText == "QRCOUNT"))
                    {
                        //整數數值欄位
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                        vFieldValue = (drTemp[i].ToString().Trim() != "") ? drTemp[i].ToString().Trim() : "0";
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(Int32.Parse(vFieldValue));
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                    }
                    else
                    {
                        vFieldValue = drTemp[i].ToString();
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.TryParse(vFieldValue, out vTempFloat) ? vTempFloat : 0.0);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Float;
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
                    string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                         "ImportTLD_Data.aspx" + Environment.NewLine +
                                         "清分日期" + fClearDate.ToString("yyyy/MM/dd");
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    //先判斷是不是IE
                    HttpContext.Current.Response.ContentType = "application/octet-stream";
                    //HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
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