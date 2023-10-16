using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Web.Configuration;

namespace Amaterasu_Function
{
    public class PublicFunction
    {
        public PublicFunction()
        {
            //
            // TODO: 在這裡新增建構函式邏輯
            //
        }

        /// <summary>
        /// 根據指定的 SQL 語法取回資料回傳
        /// </summary>
        /// <param name="fConnStr">連結字串</param>
        /// <param name="fSQLStr">SQL 查詢語法</param>
        /// <param name="fReturnField">要回傳的欄位名稱</param>
        /// <returns></returns>
        public string GetValue(string fConnStr, string fSQLStr, string fReturnField)
        {
            string vReturnStr = "";
            using (SqlConnection connTemp = new SqlConnection(fConnStr))
            {
                connTemp.Open();
                SqlCommand cmdTemp = new SqlCommand(fSQLStr, connTemp);
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                while (drTemp.Read())
                {
                    vReturnStr = drTemp[fReturnField].ToString();
                }
                cmdTemp.Cancel();
                drTemp.Close();
            }
            return vReturnStr ?? "";
        }

        /// <summary>
        /// 執行指定的 SQL 語法
        /// </summary>
        /// <param name="fConnStr">資料庫連結字串</param>
        /// <param name="fExecSQLStr">要執行的 SQL 語法</param>
        /// <returns></returns>
        public int ExecSQL(string fConnStr, string fExecSQLStr)
        {
            int vReturnStatus = 0;
            using (SqlConnection connTemp = new SqlConnection(fConnStr))
            {
                connTemp.Open();
                SqlCommand cmdTemp = new SqlCommand(fExecSQLStr, connTemp);
                vReturnStatus = cmdTemp.ExecuteNonQuery();
                cmdTemp.Cancel();
            }
            return vReturnStatus;
        }

        /// <summary>
        /// 取得 SQL 連結字串並回傳
        /// </summary>
        /// <param name="fPath">傳入網站的虛擬目錄</param>
        /// <returns>回傳存放在 web.config 裡的 SQL 連結字串</returns>
        public string GetConnectionStr(string fPath)
        {
            Configuration rootWebConfig = WebConfigurationManager.OpenWebConfiguration(fPath);
            if (0 < rootWebConfig.ConnectionStrings.ConnectionStrings.Count)
            {
                return rootWebConfig.ConnectionStrings.ConnectionStrings["connERPSQL"].ConnectionString;
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// 輸入一個日期字串，轉換成西元年字串
        /// </summary>
        /// <param name="vInputDate">要轉換的字串</param>
        /// <returns>回傳西元年字串</returns>
        public string GetAD(string vInputDate)
        {
            string vReturnStr = "";
            int vPosIndex = vInputDate.IndexOf("/");
            int vYear = Int32.Parse(vInputDate.Substring(0, vPosIndex));
            string vMonthDay = vInputDate.Substring(vPosIndex);
            vReturnStr = (vYear < 1911) ? (vYear + 1911).ToString("D4") + vMonthDay : vYear.ToString("D4") + vMonthDay;
            return vReturnStr;
        }

        /// <summary>
        /// 輸入一個西元年日期回傳民國年日期
        /// </summary>
        /// <param name="vDate">要轉換的日期</param>
        /// <returns>回傳民國年日期</returns>
        public string GetChinsesDate(DateTime vDate)
        {
            string vReturnStr = "";
            vReturnStr = (vDate.Year > 1911) ? (vDate.Year - 1911).ToString("D3") + "/" + vDate.ToString("MM/dd") : vDate.ToString("yyyy/MM/dd");
            return vReturnStr;
        }

        /// <summary>
        /// 輸入一個日期回傳同一年度的第一天
        /// </summary>
        /// <param name="vNow">要計算的日期</param>
        /// <param name="vYearType">決定回傳西元年 (B) 還是民國年 (C) </param>
        /// <returns></returns>
        public string GetYearFirstDay(DateTime vNow, string vYearType)
        {
            string vReturnStr = "";
            DateTime vThisYearFirstDay = new DateTime(vNow.Year, 1, 1);
            vReturnStr = (vYearType == "C") ?
                         (vThisYearFirstDay.Year - 1911).ToString("D3") + "/" + vThisYearFirstDay.ToString("MM/dd") : //vYearType = "C" 表示回傳民國年
                         vThisYearFirstDay.Year.ToString("D4") + "/" + vThisYearFirstDay.ToString("MM/dd"); // vYearType 不是 "C" 就回傳西元年
            return vReturnStr;
        }

        /// <summary>
        /// 輸入一個日期回傳當年最後一天
        /// </summary>
        /// <param name="vNow"></param>
        /// <param name="vYearType">決定回傳西元年 (B) 還是民國年 (C) </param>
        /// <returns></returns>
        public string GetYearLastDay(DateTime vNow, string vYearType)
        {
            string vReturnStr = "";
            DateTime vThisYearLastDay = new DateTime(vNow.Year, 12, 31);
            vReturnStr = (vYearType == "C") ?
                         (vThisYearLastDay.Year - 1911).ToString("D3") + "/" + vThisYearLastDay.ToString("MM/dd") :
                         vThisYearLastDay.Year.ToString("D4") + "/" + vThisYearLastDay.ToString("MM/dd");
            return vReturnStr;
        }

        /// <summary>
        /// 輸入一個日期回傳當月第一天
        /// </summary>
        /// <param name="vNow"></param>
        /// <param name="vYearType">決定回傳西元年 (B) 還是民國年 (C) </param>
        /// <returns></returns>
        public string GetMonthFirstDay(DateTime vNow, string vYearType)
        {
            string vReturnStr = "";
            int vThisYear = (vNow.Year > 3822) ? vNow.Year - 1911 : vNow.Year;
            DateTime vThisMonthFirstDay = new DateTime(vThisYear, vNow.Month, 1);
            vReturnStr = (vYearType == "C") ?
                         (vThisMonthFirstDay.Year - 1911).ToString("D3") + "/" + vThisMonthFirstDay.ToString("MM/dd") :
                         vThisMonthFirstDay.Year.ToString("D4") + "/" + vThisMonthFirstDay.ToString("MM/dd");
            return vReturnStr;
        }

        /// <summary>
        /// 輸入一個日期取回當月最後一天
        /// </summary>
        /// <param name="vNow"></param>
        /// <param name="vYearType">決定回傳西元年 (B) 還是民國年 (C) </param>
        /// <returns></returns>
        public string GetMonthLastDay(DateTime vNow, string vYearType)
        {
            string vReturnStr = "";
            int vNextMonth = (vNow.Month + 1 <= 12) ? vNow.Month + 1 : 1;
            int vNextMonthYear = (vNow.Month + 1 <= 12) ? vNow.Year : vNow.Year + 1;
            vNextMonthYear = (vNextMonthYear > 3822) ? vNextMonthYear - 1911 : vNextMonthYear;
            DateTime vNextMonthFirstDay = new DateTime(vNextMonthYear, vNextMonth, 1);
            DateTime vThisMonthLastDay = vNextMonthFirstDay.AddDays(-1);
            vReturnStr = (vYearType == "C") ?
                         (vThisMonthLastDay.Year - 1911).ToString("D3") + "/" + vThisMonthLastDay.ToString("MM/dd") :
                         vThisMonthLastDay.Year.ToString("D4") + "/" + vThisMonthLastDay.ToString("MM/dd");
            return vReturnStr;
        }

        /// <summary>
        /// 輸入一個日期字串轉換成指定格式的字串
        /// </summary>
        /// <param name="SourceDateString">要轉換的日期字串</param>
        /// <param name="vYearType">決定回傳西元年 (B) 還是民國年 (C) 或是加國號的民國年 (C2)</param>
        /// <returns></returns>
        public string TransDateString(string SourceDateString, string vYearType)
        {
            string vReturnStr = "";
            DateTime vTempDate = DateTime.Parse(SourceDateString);
            if (vTempDate.Year > 3822)
            {
                vTempDate = vTempDate.AddYears(-1911);
            }
            vReturnStr = (vYearType == "C") ?
                         (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd") :
                         (vYearType == "C2") ?
                         "中華民國 " + (vTempDate.Year - 1911).ToString("D3") + " 年 " + vTempDate.Month.ToString("D2") + " 月 " + vTempDate.Day.ToString("D2") + " 日" :
                         (vYearType == "B") ?
                         vTempDate.Year.ToString("D4") + "/" + vTempDate.ToString("MM/dd") : "";
            return vReturnStr;
        }

        /// <summary>
        /// 輸入一個日期，回傳指定格式的字串
        /// </summary>
        /// <param name="SourceDate">要轉換的日期</param>
        /// <param name="vYearType">決定回傳西元年 (B) 還是民國年 (C) </param>
        /// <returns></returns>
        public string TransDateString(DateTime SourceDate, string vYearType)
        {
            string vReturnStr = "";
            DateTime vTempDate = ((SourceDate.Year > 1911) && (SourceDate.Year > 3822)) ? SourceDate.AddYears(-1911) :
                                 (SourceDate.Year < 1911) ? SourceDate.AddYears(1911) : SourceDate;
            vReturnStr = (vYearType == "C") ?
                         (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd") :
                         vTempDate.Year.ToString("D4") + "/" + vTempDate.ToString("MM/dd");
            return vReturnStr;
        }

        /// <summary>
        /// 輸入一個日期取回當月份天數
        /// </summary>
        /// <param name="vNow">要計算天數的日期</param>
        /// <returns>天數</returns>
        public int GetMonthDays(DateTime vNow)
        {
            int vResultInt = 0;
            int vNextMonth = (vNow.Month + 1 <= 12) ? vNow.Month + 1 : 1;
            DateTime vNextMonthFirstDay = new DateTime(vNow.Year, vNextMonth, 1);
            DateTime vThisMonthLastDay = vNextMonthFirstDay.AddDays(-1);
            vResultInt = vThisMonthLastDay.Day;
            return vResultInt;
        }

        /// <summary>
        /// 取得指定員工在指定的日期範圍內的工作天數
        /// </summary>
        /// <param name="fConnStr">SQL 連接字串</param>
        /// <param name="fEmpNo">員工編號</param>
        /// <param name="fFirstDate">計算開始日期</param>
        /// <param name="fLastDate">計算結束日期</param>
        /// <returns>回傳工作天數 (整數)</returns>
        public int CalEmployeeWorkDays(string fConnStr, string fEmpNo, DateTime fFirstDate, DateTime fLastDate)
        {
            int vResultDays = 0;
            string vFirstDate = fFirstDate.Year.ToString("D4") + "/" + fFirstDate.ToString("MM/dd");
            string vLastDate = fLastDate.Year.ToString("D4") + "/" + fLastDate.ToString("MM/dd");
            using (SqlConnection connTemp = new SqlConnection(fConnStr))
            {
                string vTempSQLStr = "select MonthDays from (" + Environment.NewLine +
                                     "                       select case when (AssumeDay > '" + vLastDate + "') or (LeaveDay - 1 < '" + vFirstDate + "') then 0 " + Environment.NewLine + //在計薪年月後才到職或是計薪年月前就離職了
                                     "                                   when BeginDay - 1 < '" + vFirstDate + "' and StopDay > '" + vLastDate + "' then 0 " + Environment.NewLine +//計薪年月整個月都處於留職停薪的狀態
                                     "                                   when (BeginDay <= AssumeDay) or (StopDay < AssumeDay) then 0 " + Environment.NewLine +//留職停薪日比到職日還早，表示不正常
                                     "                                   when BeginDay is null and StopDay is null then " + Environment.NewLine +//沒有留職停薪的人
                                     "                                        DateDiff(DAY,case when AssumeDay < '" + vFirstDate + "' then '" + vFirstDate + "' else AssumeDay end," + Environment.NewLine +
                                     "                                                     case when (LeaveDay is null) or LeaveDay - 1 > '" + vLastDate + "' then '" + vLastDate + "' else LeaveDay - 1 end) + 1 " + Environment.NewLine +
                                     "                                   when BeginDay is not null and StopDay is null then " + Environment.NewLine +//有留職停薪開始日可是沒有結束日，表示還在留職停薪之中
                                     "                                        DateDiff(DAY,case when AssumeDay < '" + vFirstDate + "' then '" + vFirstDate + "' else AssumeDay end," + Environment.NewLine +
                                     "                                                     case when BeginDay - 1 < '" + vFirstDate + "' then '" + vFirstDate + "' else BeginDay end)" + Environment.NewLine +
                                     "                                   when BeginDay is not null and StopDay is not null then " + Environment.NewLine +//留職停薪起迄日期都有，表示已經復職了
                                     "                                        DateDiff(DAY,case when AssumeDay < '" + vFirstDate + "' then '" + vFirstDate + "' else AssumeDay end," + Environment.NewLine +
                                     "                                                     case when BeginDay - 1 < '" + vFirstDate + "' then '" + vFirstDate + "' " + Environment.NewLine +
                                     "                                                          when BeginDay - 1 > '" + vLastDate + "' then '" + vLastDate + "' " + Environment.NewLine +
                                     "                                                          else BeginDay end) + " + Environment.NewLine +
                                     "                                        DateDiff(DAY,case when isnull(ReturnDay,StopDay) between '" + vFirstDate + "' and '" + vLastDate + "' then isnull(ReturnDay,StopDay) " + Environment.NewLine +
                                     "                                                          when isnull(ReturnDay,StopDay) < '" + vFirstDate + "' then '" + vFirstDate + "' " + Environment.NewLine +
                                     "                                                          else '" + vLastDate + "' end," + Environment.NewLine +
                                     "                                                     case when isnull(LeaveDay - 1,'" + vLastDate + "') between '" + vFirstDate + "' and '" + vLastDate + "' then isnull(LeaveDay - 1,'" + vLastDate + "') " + Environment.NewLine +
                                     "                                                          else '" + vLastDate + "' end) + 1 " + Environment.NewLine +
                                     "                                   when BeginDay is null and StopDay is not null then " + Environment.NewLine +//沒有開始留停日卻有復職日...把復職日當做到職日計算
                                     "                                        DateDiff(DAY,case when isnull(ReturnDay,StopDay) < '" + vFirstDate + "' then '" + vFirstDate + "' else isnull(ReturnDay,StopDay) end," + Environment.NewLine +
                                     "                                                     case when (LeaveDay is null) or LeaveDay > '" + vLastDate + "' then '" + vLastDate + "' else LeaveDay end) + 1 " + Environment.NewLine +
                                     "                              end as MonthDays, DepNo, EmpNo, Title, AssumeDay, LeaveDay, ReturnDay, BeginDay, StopDay from Employee " + Environment.NewLine +
                                     "                        where EmpNo = '" + fEmpNo + "' " + Environment.NewLine +
                                     "                      ) a ";
                SqlCommand cmdTemp = new SqlCommand(vTempSQLStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                while (drTemp.Read())
                {
                    vResultDays = int.Parse(drTemp["MonthDays"].ToString().Trim());
                }
                cmdTemp.Cancel();
                drTemp.Close();
            }
            return vResultDays;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fConnStr"></param>
        /// <param name="fEmpNo"></param>
        /// <param name="fCalYM"></param>
        /// <returns></returns>
        public int CalRetireWorkDays(string fConnStr, string fEmpNo, DateTime fCalYM)
        {
            int vResultDays = 0;
            int vYear_New = (fCalYM.Month + 1 > 12) ? fCalYM.Year + 1 : fCalYM.Year;
            int vMonth_New = (fCalYM.Month + 1 > 12) ? 1 : fCalYM.Month + 1;
            DateTime vEmpArriveDate = DateTime.FromBinary(0);
            DateTime vEmpLeaveDate = DateTime.FromBinary(0);
            DateTime vMonthFirstDate = DateTime.Parse(fCalYM.Year.ToString("D4") + "/" + fCalYM.ToString("MM") + "/01");
            DateTime vMonthLastDate = DateTime.Parse(vYear_New.ToString("D4") + "/" + vMonth_New.ToString("D2") + "/01").AddDays(-1);

            SqlConnection connTemp = new SqlConnection(fConnStr);
            string vSQLStr = "select isnull(convert(char(10), AssumeDay, 111), 0) AssumeDay, " + Environment.NewLine +
                             "       isnull(convert(char(10), LeaveDay, 111), 0) LeaveDay " + Environment.NewLine +
                             "  from Employee where EmpNo = '" + fEmpNo + "' ";
            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
            connTemp.Open();
            SqlDataReader drTemp = cmdTemp.ExecuteReader();
            while (drTemp.Read())
            {
                vEmpArriveDate = (drTemp["AssumeDay"].ToString().Trim() != "0") ? DateTime.Parse(drTemp["AssumeDay"].ToString().Trim()) : vMonthFirstDate;
                vEmpLeaveDate = (drTemp["LeaveDay"].ToString().Trim() != "0") ? DateTime.Parse(drTemp["LeaveDay"].ToString().Trim()) : vMonthLastDate;
            }

            return vResultDays;
        }

        /// <summary>
        /// 時間字串轉分鐘數
        /// </summary>
        /// <param name="fTimeString"></param>
        /// <returns></returns>
        public int TimeStrToMins(string fTimeString)
        {
            int vResultMins = 0;
            int vSplitPOS = fTimeString.IndexOf(":");
            int vHours = Int32.Parse(fTimeString.Substring(0, vSplitPOS));
            int vMins = Int32.Parse(fTimeString.Substring(vSplitPOS + 1, fTimeString.Length - (vSplitPOS + 1)));
            vResultMins = vHours * 60 + vMins;
            return vResultMins;
        }

        /// <summary>
        /// 取得流水序號
        /// </summary>
        /// <param name="fConnStr">資料庫連接字串</param>
        /// <param name="fTableName">資料表名稱</param>
        /// <param name="fFieldName">序號欄位名稱</param>
        /// <param name="fYearType">年份格式 'C':民國年 'B':西元年</param>
        /// <param name="fYMStrHasDays">日期格式是否包括日</param>
        /// <param name="fCodeDate">用來取號的日期</param>
        /// <param name="fYMLength">日期部份長度</param>
        /// <param name="fIndexCode">序號中置碼</param>
        /// <param name="fIndexLength">序號總長度</param>
        /// <returns></returns>
        public string GetDataIndex(string fConnStr, string fTableName, string fFieldName, string fYearType, Boolean fYMStrHasDays, DateTime fCodeDate, int fYMLength, string fIndexCode, int fIndexLength)
        {
            string vResultStr = "";
            string vFirstCode = ((fYearType == "C") && (fYMStrHasDays == true)) ? (fCodeDate.Year - 1911).ToString("D3") + fCodeDate.ToString("MMdd") :
                                ((fYearType == "C") && (fYMStrHasDays == false)) ? (fCodeDate.Year - 1911).ToString("D3") + fCodeDate.ToString("MM") :
                                ((fYearType == "B") && (fYMStrHasDays == true)) ? fCodeDate.Year.ToString("D4") + fCodeDate.ToString("MMdd") :
                                ((fYearType == "B") && (fYMStrHasDays == false)) ? fCodeDate.Year.ToString("D4") + fCodeDate.ToString("MM") : "";
            vFirstCode = ("0000" + vFirstCode).Substring((vFirstCode.Length + 4) - fYMLength);
            string vSQLStr_Temp = "select Max(" + fFieldName.Trim() + ") MaxNo from " + fTableName.Trim() + " where " + fFieldName.Trim() + " like '" + vFirstCode.Trim() + fIndexCode.Trim() + "%'";
            string vMaxNo = "";
            int vIndex = 0;
            int vIndexLen = 0;
            string vIndexStr = "";
            using (SqlConnection connTemp = new SqlConnection(fConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                while (drTemp.Read())
                {
                    vMaxNo = drTemp["MaxNo"].ToString().Trim();
                }
                cmdTemp.Cancel();
                drTemp.Close();
            }
            vIndex = (vMaxNo == "") ? 1 : Int32.Parse(vMaxNo.Replace(vFirstCode + fIndexCode, "")) + 1;
            vIndexStr = vIndex.ToString().Trim();
            vIndexLen = vIndexStr.Length;
            for (int i = 0; i < (fIndexLength - vIndexLen); i++)
            {
                vIndexStr = "0" + vIndexStr;
            }
            vResultStr = vFirstCode.Trim() + fIndexCode.Trim() + vIndexStr.Trim();
            return vResultStr;
        }

        /// <summary>
        /// 檢查日期是否已經關帳
        /// </summary>
        /// <param name="fConnStr">資料庫連接字串</param>
        /// <param name="fControlItem">關帳日期欄位識別碼</param>
        /// <param name="fContent">要比對的日期</param>
        /// <returns></returns>
        public bool DateIsClosed(string fConnStr, string fControlItem, string fContent)
        {
            bool vResultValue = false;
            string vSQLStr_Temp = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = '" + fControlItem.Trim() + "' ";
            string vContent_Old = "";
            using (SqlConnection connTemp = new SqlConnection(fConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                while (drTemp.Read())
                {
                    vContent_Old = drTemp["Content"].ToString().Trim();
                }
                cmdTemp.Cancel();
                drTemp.Close();
            }
            vResultValue = (vContent_Old == "") ? false : (Int32.Parse(vContent_Old) >= Int32.Parse(fContent.Trim()));
            return vResultValue;
        }

        /// <summary>
        /// 把沒有分隔字元的日期字串轉換成日期格式
        /// </summary>
        /// <param name="DateStr">沒有分隔字元的日期字串</param>
        /// <param name="DateType">"C"表示回傳民國年字串，"B"表示回傳西元年字串</param>
        /// <returns>轉換成功時回傳日期字串，失敗時回傳 '0'</returns>
        public string TransNoneSymbolDateStrToDate(string DateStr, string DateType)
        {
            string vResultDate;
            Int32 vDateNumber = 0;
            string vDateStr_Source = DateStr.Replace("/", "");
            vDateStr_Source = vDateStr_Source.Replace(".", "");
            string vDateStr_BC = "";
            string vDate_Y = "";
            string vDate_M = "";
            string vDate_D = "";

            if ((vDateStr_Source.Length == 7) && (Int32.TryParse(vDateStr_Source, out vDateNumber)))
            {
                switch (DateType)
                {
                    case "B":
                        vDateNumber = Int32.Parse(vDateStr_Source) + 19110000;
                        vDateStr_BC = vDateNumber.ToString();
                        vDate_Y = vDateStr_BC.Substring(0, 4);
                        vDate_M = vDateStr_BC.Substring(4, 2);
                        vDate_D = vDateStr_BC.Substring(6, 2);
                        break;
                    case "C":
                        vDate_Y = DateStr.Substring(0, 3);
                        vDate_M = DateStr.Substring(3, 2);
                        vDate_D = DateStr.Substring(5, 2);
                        break;
                }
                vResultDate = vDate_Y + "/" + vDate_M + "/" + vDate_D;
            }
            else if ((vDateStr_Source.Length == 8) && (Int32.TryParse(vDateStr_Source, out vDateNumber)))
            {
                switch (DateType)
                {
                    case "C":
                        vDateNumber = Int32.Parse(vDateStr_Source) - 19110000;
                        vDateStr_BC = vDateNumber.ToString();
                        vDate_Y = vDateStr_BC.Substring(0, 3);
                        vDate_M = vDateStr_BC.Substring(3, 2);
                        vDate_D = vDateStr_BC.Substring(5, 2);
                        break;
                    case "B":
                        vDate_Y = DateStr.Substring(0, 4);
                        vDate_M = DateStr.Substring(4, 2);
                        vDate_D = DateStr.Substring(6, 2);
                        break;
                }
                vResultDate = vDate_Y + "/" + vDate_M + "/" + vDate_D;
            }
            else
            {
                vResultDate = "0";
            }
            return vResultDate;
        }

        /// <summary>
        /// 新增操作記錄
        /// </summary>
        /// <param name="fConnStr">資料庫連結字串</param>
        /// <param name="fLoginID">登入帳號</param>
        /// <param name="fComputerName">電腦名稱</param>
        /// <param name="fOperationNote">操作內容</param>
        /// <returns></returns>
        public int InsertOperateRecord(string fConnStr, string fLoginID, string fComputerName, string fOperationNote)
        {
            int vResultInt = 0;
            string vSQLStr = "insert into OperateRecord(LoginID, ComputerName, OperateDate, OperationNote)" + Environment.NewLine +
                             " values('" + fLoginID + "', '" + fComputerName + "', GetDate(), '" + fOperationNote + "')";
            using (SqlConnection connTemp = new SqlConnection(fConnStr))
            {
                connTemp.Open();
                SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                vResultInt = cmdTemp.ExecuteNonQuery();
                cmdTemp.Cancel();
            }
            return vResultInt;
        }

        /// <summary>
        /// 輸入國字大寫數字或全形數字字串轉出阿拉伯數字字串
        /// </summary>
        /// <param name="fChineseNumber"></param>
        /// <returns></returns>
        public string ChangeChineseNumbers(string fChineseNumber)
        {
            string vResultStr = "";
            int TempInt = 0;
            string[] vCNumberList_1 = { "Ｏ", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
            string[] vCNumberList_2 = { "０", "１", "２", "３", "４", "５", "６", "７", "８", "９" };
            string vTempChar = "";
            string vTempResultChar = "";
            for (int i = 0; i < fChineseNumber.Length; i++)
            {
                vTempChar = fChineseNumber.Substring(i, 1);
                if (Int32.TryParse(vTempChar, out TempInt))
                {
                    vTempResultChar = vTempChar;
                }
                else if (vTempChar.Trim() == "○")
                {
                    vTempResultChar = "0";
                }
                else
                {
                    for (int j = 0; j < vCNumberList_1.Length; j++)
                    {
                        vTempResultChar = (vTempChar == vCNumberList_1[j]) ? j.ToString() : vTempResultChar;
                    }
                    for (int j = 0; j < vCNumberList_2.Length; j++)
                    {
                        vTempResultChar = (vTempChar == vCNumberList_2[j]) ? j.ToString() : vTempResultChar;
                    }
                }
                vResultStr += vTempResultChar;
            }
            return vResultStr;
        }

        /// <summary>
        /// 輸入整數字串輸出國字數字
        /// </summary>
        /// <param name="fOriNumber"></param>
        /// <returns></returns>
        public string GetChineseNumber(string fOriNumber)
        {
            string vResultStr = "";
            string[] vCNumberList = { "Ｏ", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
            string vTempChar = "";
            int vCharIndex = 0;
            for (int i = 0; i < fOriNumber.Length; i++)
            {
                vTempChar = fOriNumber.Substring(i, 1);
                if (Int32.TryParse(vTempChar, out vCharIndex))
                {
                    vResultStr += vCNumberList[Int32.Parse(vTempChar)];
                }
            }
            return vResultStr;
        }

        /// <summary>
        /// 輸入整數字串輸出全形字串
        /// </summary>
        /// <param name="fOriNumber"></param>
        /// <returns></returns>
        public string GetFullcharNumber(string fOriNumber)
        {
            string vResultStr = "";
            string[] vCNumberList = { "０", "１", "２", "３", "４", "５", "６", "７", "８", "９" };
            string vTempChar = "";
            int vCharIndex = 0;
            for (int i = 0; i < fOriNumber.Length; i++)
            {
                vTempChar = fOriNumber.Substring(i, 1);
                if (Int32.TryParse(vTempChar, out vCharIndex))
                {
                    vResultStr += vCNumberList[Int32.Parse(vTempChar)];
                }
            }
            return vResultStr;
        }

        /// <summary>
        /// 計算
        /// </summary>
        /// <param name="fConsNo"></param>
        /// <param name="fConnStr"></param>
        public void CalIAStoreQuantity(string fConsNo, string fConnStr, string fSheetNo)
        {
            //取回一般進出貨單的數據
            string vTempStr = (fSheetNo == "") ? 
                              "select sum(Quantity * QtyMode) TQty from IASheetB where ItemStatus in ('001', '102') and ConsNo = '" + fConsNo + "' " :
                              "select sum(Quantity * QtyMode) TQty from IASheetB where ItemStatus in ('001', '102') and ConsNo = '" + fConsNo + "' and SheetNo != '" + fSheetNo.Trim() + "' ";
            string vTQtyStr = GetValue(fConnStr, vTempStr, "TQty");
            int vTempQty = 0;
            int vTQty = 0;
            if (!Int32.TryParse(vTQtyStr,out vTempQty))
            {
                vTQty += vTempQty;
            }
            else
            {
                vTQty += 0;
            }
            //取回盤點單的數據
            vTempStr = "select sum(Quantity * QtyMode) TQty from IASheetB where ConsNo = '" + fConsNo + "' and SheetNo in (select SheetNo from IASheetA where SheetMode = 'SS') ";
            vTQtyStr = GetValue(fConnStr, vTempStr, "TQty");
            if (!Int32.TryParse(vTQtyStr, out vTempQty))
            {
                vTQty += vTempQty;
            }
            else
            {
                vTQty += 0;
            }
            vTempStr = "update IAConsumables set Quantity =" + vTQty.ToString() + " where ConsNo = '" + fConsNo + "' ";
            ExecSQL(fConnStr, vTempStr);
        }

        /// <summary>
        /// 取回耗材進出貨單號
        /// </summary>
        /// <param name="fConnStr"></param>
        /// <param name="fSheetMode"></param>
        /// <returns></returns>
        public string GetIASheetNo(string fConnStr, string fSheetMode)
        {
            string vResultStr = "";
            string vSQLStr_Temp = "select max(SheetNo) MaxIndex from IASheetA where SheetMode = '" + fSheetMode + "' ";
            string vOldSheetNo = GetValue(fConnStr, vSQLStr_Temp, "MaxIndex");
            string vOldIndex = (vOldSheetNo != "") ? vOldSheetNo.Substring(8) : "0";
            int vNewIndex = 0;
            if (int.TryParse(vOldIndex, out vNewIndex))
            {
                vNewIndex++;
            }
            else
            {
                vNewIndex = 1;
            }
            vResultStr = DateTime.Now.Year.ToString("D4") + DateTime.Now.Month.ToString("D2") + fSheetMode + vNewIndex.ToString("D4");
            return vResultStr.Trim();
        }
    }
}