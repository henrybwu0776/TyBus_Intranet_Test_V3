using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace TyBus_Intranet_Test_V3
{
    public partial class TicketTrans : Page
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
                vLoginID = (Session["LoginID"] != null) ? Session["LoginID"].ToString().Trim() : "";
                vLoginName = (Session["LoginName"] != null) ? Session["LoginName"].ToString().Trim() : "";
                vLoginDepNo = (Session["LoginDepNo"] != null) ? Session["LoginDepNo"].ToString().Trim() : "";
                vLoginDepName = (Session["LoginDepName"] != null) ? Session["LoginDepName"].ToString().Trim() : "";
                vLoginTitle = (Session["LoginTitle"] != null) ? Session["LoginTitle"].ToString().Trim() : "";
                vLoginTitleName = (Session["LoginTitleName"] != null) ? Session["LoginTitleName"].ToString().Trim() : "";
                vLoginEmpType = (Session["LoginEmpType"] != null) ? Session["LoginEmpType"].ToString().Trim() : "";
                //vComputerName = Environment.MachineName; //2021.09.27 新增取得電腦名稱
                vComputerName = Page.Request.UserHostName; //2021.10.08 改變取得電腦名稱的方法

                UnobtrusiveValidationMode = System.Web.UI.UnobtrusiveValidationMode.None;

                if (vLoginID != "")
                {
                    //事件日期
                    string vBuDateURL = "InputDate.aspx?TextboxID=" + eBuDate_S_Search.ClientID;
                    string vBuDateScript = "window.open('" + vBuDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate_S_Search.Attributes["onClick"] = vBuDateScript;

                    vBuDateURL = "InputDate.aspx?TextboxID=" + eBuDate_E_Search.ClientID;
                    vBuDateScript = "window.open('" + vBuDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate_E_Search.Attributes["onClick"] = vBuDateScript;

                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {

                    }
                    else
                    {
                        OpenData(eBuDate_S_Search.Text.Trim(), eBuDate_E_Search.Text.Trim());
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

        private void OpenData(string fDate_S, string fDate_E)
        {
            string vSelectStr = "SELECT CONVERT (char(8), budate, 112) + CONVERT (char(8), date, 112) + CAST(depno AS char(12)) + CAST(linesno AS char(12)) + CAST(car_id AS char(12)) + CAST(car_no AS char(12)) + CAST(driver AS char(12)) AS IndexNo, " + Environment.NewLine +
                                "       budate, date, depno, linesno, car_id, car_no, driver " + Environment.NewLine +
                                "  FROM CPSOUTG " + Environment.NewLine +
                                " WHERE isnull(BuDate, '') <> '' ";
            string vWStr_BuDate = "";
            if ((fDate_S.Trim() != "") && (fDate_E.Trim() != ""))
            {
                vWStr_BuDate = "   AND BuDate between '" + fDate_S.Trim() + "' and '" + fDate_E.Trim() + "' ";
            }
            else if ((fDate_S.Trim() != "") && (fDate_E.Trim() == ""))
            {
                vWStr_BuDate = "   AND BuDate = '" + fDate_S.Trim() + "' ";
            }
            else if ((fDate_S.Trim() == "") && (fDate_E.Trim() != ""))
            {
                vWStr_BuDate = "   AND BuDate = '" + fDate_E.Trim() + "' ";
            }

            if (vWStr_BuDate != "")
            {
                sdsCPSOutG_List.SelectCommand = vSelectStr + Environment.NewLine +
                                                vWStr_BuDate + Environment.NewLine +
                                                " Order by BuDate, Date, DepNo, LinesNo, Car_ID";
                gridCPSOutG.DataBind();
                fvCPSOutG.DataBind();
            }
        }

        /// <summary>
        /// 查詢清分資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Search_Click(object sender, EventArgs e)
        {
            if ((eBuDate_S_Search.Text.Trim() == "") && (eBuDate_E_Search.Text.Trim() == ""))
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先輸入要查詢的日期區間')");
                Response.Write("</" + "Script>");
            }
            else
            {
                OpenData(eBuDate_S_Search.Text.Trim(), eBuDate_E_Search.Text.Trim());
            }
        }

        /// <summary>
        /// 匯入清分資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbImport_Search_Click(object sender, EventArgs e)
        {
            if (fuTicketTrans.FileName.Trim() != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuTicketTrans.FileName); //取回副檔名
                string vSQLStrTemp = "";
                string vErrorStr = "";
                int vIndex = 0;

                string vSourceName = "";
                DateTime vGetDate_Temp;
                string vBuDateStr_In = "";
                string vDate_In = "";
                string vCalYM_In = "";
                string vTicketType_In = "01"; //固定只轉一般卡的票證收入資料
                string vDepNo_In = "";
                string vLinesNo_In = "";
                string vTicketLinesNo_In = "";
                string vGovLinesNo_In = "";
                string vCarID_In = "";
                string vCarNo_In = "";
                string vDriver_In = "";
                double vOAmt_In = 0.0;
                double vOPiece_In = 0.0;
                double vSPiece_In = 0.0;
                double vOldPiece_In = 0.0;
                int vHCount_In = 0;
                int vPCount_In = 0;
                int vLCount_In = 0;
                int vSCount_In = 0;
                int vOldCount_In = 0;
                string vIndexNo_In = "";

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuTicketTrans.FileContent);
                        for (int si = 0; si < wbExcel_H.NumberOfSheets; si++)
                        {
                            HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(si);
                            if (sheetExcel_H.SheetName.Trim() == "清分明細報表")
                            {
                                //先進行一次資料檢查
                                for (int i = 0; i <= sheetExcel_H.LastRowNum; i++)
                                {
                                    //逐行讀入資料
                                    HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                                    if ((vRowTemp_H != null) && (vRowTemp_H.Cells.Count >= 16))
                                    {
                                        vSourceName = vRowTemp_H.Cells[0].ToString().Trim();
                                        if ((vSourceName != "") && (vSourceName != "票證公司"))
                                        {
                                            vBuDateStr_In = vRowTemp_H.Cells[1].ToString();
                                            if (DateTime.TryParse(vBuDateStr_In, out vGetDate_Temp))
                                            {
                                                vBuDateStr_In = ((vSourceName == "悠遊卡") || (vSourceName == "一卡通")) ? vGetDate_Temp.AddDays(1).ToString("yyyy/MM/dd") : vGetDate_Temp.ToString("yyyy/MM/dd");
                                            }
                                            else
                                            {
                                                vErrorStr += (Environment.NewLine + "'清分日期有誤，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            vDate_In = vRowTemp_H.Cells[2].ToString();
                                            if (DateTime.TryParse(vDate_In, out vGetDate_Temp))
                                            {
                                                vDate_In = vGetDate_Temp.ToString("yyyy/MM/dd");
                                            }
                                            else
                                            {
                                                vErrorStr += (Environment.NewLine + "'營運日期有誤，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            vDepNo_In = vRowTemp_H.Cells[3].ToString();
                                            vSQLStrTemp = "select isnull([Name], '') as [Name] from Department where DepNo = '" + vDepNo_In + "' ";
                                            if (PF.GetValue(vConnStr, vSQLStrTemp, "Name") == "")
                                            {
                                                vErrorStr += (Environment.NewLine + "'站別代碼有誤，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            vCarID_In = vRowTemp_H.Cells[4].ToString();
                                            if ((vCarID_In.Substring(0, 1) != "4") && (vCarID_In.IndexOf("-U7") > 0))
                                            {
                                                vCarID_In.Replace("-U7", "-U-7");
                                            }
                                            vSQLStrTemp = "select Car_No from Car_InfoA where Car_ID = '" + vCarID_In + "' ";
                                            if (PF.GetValue(vConnStr, vSQLStrTemp, "Car_No") == "")
                                            {
                                                vErrorStr += (Environment.NewLine + "'牌照號碼有誤，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            vDriver_In = Int32.Parse(vRowTemp_H.Cells[5].ToString()).ToString("D6");
                                            vSQLStrTemp = "select [Name] from Employee where EmpNo = '" + vDriver_In + "' ";
                                            if (PF.GetValue(vConnStr, vSQLStrTemp, "Name") == "")
                                            {
                                                vErrorStr += (Environment.NewLine + "'駕駛員工號有誤，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            vTicketLinesNo_In = vRowTemp_H.Cells[6].ToString();
                                            vSQLStrTemp = "select ERPLinesNo from LinesNoChart where TicketLineNo = '" + Int32.Parse(vTicketLinesNo_In).ToString() + "' ";
                                            if (PF.GetValue(vConnStr, vSQLStrTemp, "ERPLinesNo") == "")
                                            {
                                                vErrorStr += (Environment.NewLine + "'驗票機路線代碼有誤，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!double.TryParse(vRowTemp_H.Cells[8].ToString(), out vOAmt_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'營運金額必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            //2022/04/12 詢問主計課業管人員，目前票證公司提供的檔案中已經沒有聯營卡點數資料，聯營卡實際上也已經沒有使用了，所以固定帶 0
                                            //if (!double.TryParse(vRowTemp_H.Cells[8].ToString(), out vOPiece_In))
                                            //{
                                            //    vErrorStr += (Environment.NewLine + "'聯營卡點數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            //}
                                            if (!double.TryParse(vRowTemp_H.Cells[11].ToString(), out vSPiece_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'學生卡點數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!double.TryParse(vRowTemp_H.Cells[15].ToString(), out vOldPiece_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'敬老點數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!Int32.TryParse(vRowTemp_H.Cells[20].ToString(), out vHCount_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'總載客人數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!Int32.TryParse(vRowTemp_H.Cells[21].ToString(), out vPCount_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'普卡人數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!Int32.TryParse(vRowTemp_H.Cells[22].ToString(), out vLCount_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'聯營卡人數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!Int32.TryParse(vRowTemp_H.Cells[25].ToString(), out vSCount_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'學生卡人數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!Int32.TryParse(vRowTemp_H.Cells[28].ToString(), out vOldCount_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'敬老人數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        vErrorStr += (Environment.NewLine + "'EXCEL 檔案格式有誤!' ");
                                    }
                                }
                                if (vErrorStr.Trim() == "")
                                {
                                    for (int i = 0; i <= sheetExcel_H.LastRowNum; i++)
                                    {
                                        //逐行讀入資料
                                        HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                                        vSourceName = vRowTemp_H.Cells[0].ToString().Trim();
                                        if ((vSourceName != "") && (vSourceName != "票證公司"))
                                        {
                                            vBuDateStr_In = vRowTemp_H.Cells[1].ToString();
                                            vGetDate_Temp = DateTime.Parse(vBuDateStr_In);
                                            vBuDateStr_In = ((vSourceName == "悠遊卡") || (vSourceName == "一卡通")) ? vGetDate_Temp.AddDays(1).ToString("yyyy/MM/dd") : vGetDate_Temp.ToString("yyyy/MM/dd");
                                            vCalYM_In = (DateTime.Parse(vBuDateStr_In).Year - 1911).ToString("D3") + (DateTime.Parse(vBuDateStr_In).Month.ToString("D2"));
                                            vDate_In = vRowTemp_H.Cells[2].ToString();
                                            vGetDate_Temp = DateTime.Parse(vDate_In);
                                            vDate_In = vGetDate_Temp.ToString("yyyy/MM/dd");
                                            vDepNo_In = vRowTemp_H.Cells[3].ToString();
                                            vCarID_In = vRowTemp_H.Cells[4].ToString();
                                            if ((vCarID_In.Substring(0, 1) != "4") && (vCarID_In.IndexOf("-U7") > 0))
                                            {
                                                vCarID_In.Replace("-U7", "-U-7");
                                            }
                                            vSQLStrTemp = "select Car_No from Car_InfoA where Car_ID = '" + vCarID_In + "' ";
                                            vCarNo_In = PF.GetValue(vConnStr, vSQLStrTemp, "Car_No");
                                            vDriver_In = Int32.Parse(vRowTemp_H.Cells[5].ToString()).ToString("D6");
                                            vTicketLinesNo_In = vRowTemp_H.Cells[6].ToString();
                                            vSQLStrTemp = "select ERPLinesNo from LinesNoChart where TicketLineNo = '" + Int32.Parse(vTicketLinesNo_In).ToString() + "' ";
                                            vLinesNo_In = PF.GetValue(vConnStr, vSQLStrTemp, "ERPLinesNo");
                                            vIndexNo_In = vCalYM_In + vTicketType_In + vTicketLinesNo_In;
                                            vOAmt_In = double.Parse(vRowTemp_H.Cells[8].ToString());
                                            //2022/04/12 詢問主計課業管人員，目前票證公司提供的檔案中已經沒有聯營卡點數資料，聯營卡實際上也已經沒有使用了，所以固定帶 0
                                            vOPiece_In = 0.0;
                                            vSPiece_In = double.Parse(vRowTemp_H.Cells[11].ToString());
                                            vOldPiece_In = double.Parse(vRowTemp_H.Cells[15].ToString());
                                            vHCount_In = Int32.Parse(vRowTemp_H.Cells[20].ToString());
                                            vPCount_In = Int32.Parse(vRowTemp_H.Cells[21].ToString());
                                            vLCount_In = Int32.Parse(vRowTemp_H.Cells[22].ToString());
                                            vSCount_In = Int32.Parse(vRowTemp_H.Cells[25].ToString());
                                            vOldCount_In = Int32.Parse(vRowTemp_H.Cells[28].ToString());
                                        }
                                        try
                                        {
                                            vSQLStr_Main = "insert into CPSOutG " + Environment.NewLine +
                                                           "            (BuDate, Date, DepNo, LinesNo, Car_ID, Car_No, Driver, OAmt, OPiece, SPiece, " + Environment.NewLine +
                                                           "             HCount, SCount, OldPiece, PCount, LCount, OldCount)" + Environment.NewLine +
                                                           "      values(@BuDate, @Date, @DepNo, @LinesNo, @Car_ID, @Car_No, @Driver, @OAmt, @OPiece, @SPiece, " + Environment.NewLine +
                                                           "             @HCount, @SCount, @OldPiece, @PCount, @LCount, @OldCount)";
                                            sdsCPSOutG_Detail.InsertCommand = vSQLStr_Main;
                                            sdsCPSOutG_Detail.InsertParameters.Clear();
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, vBuDateStr_In));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("Date", DbType.Date, vDate_In));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo_In));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("Car_ID", DbType.String, vCarID_In));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("Car_No", DbType.String, vCarNo_In));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("Driver", DbType.String, vDriver_In));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("OAmt", DbType.Double, vOAmt_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("OPiece", DbType.Double, vOPiece_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("SPiece", DbType.Double, vSPiece_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("OldPiece", DbType.Double, vOldPiece_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("HCount", DbType.Double, vHCount_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("SCount", DbType.Double, vSCount_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("PCount", DbType.Double, vPCount_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("LCount", DbType.Double, vLCount_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("OldCount", DbType.Double, vOldCount_In.ToString()));
                                            sdsCPSOutG_Detail.Insert();
                                            gridCPSOutG.DataBind();

                                            vSQLStrTemp = "select IndexNo from LinesBounds where IndexNo = '" + vIndexNo_In + "' ";
                                            if (PF.GetValue(vConnStr,vSQLStrTemp,"IndeexNo")=="")
                                            {
                                                vSQLStr_Main = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLinesNo, LinesGOVNo, TicketPCount, TicketIncome)" + Environment.NewLine +
                                                               "     values ('" + vIndexNo_In + "', '" + vCalYM_In + "', '" + vTicketType_In + "', '" + vTicketLinesNo_In + "', '" + vGovLinesNo_In + "', " + vHCount_In.ToString() + ", " + vOAmt_In.ToString() + ")";
                                            }
                                            else
                                            {
                                                vSQLStr_Main = "update LinesBounds set TicketPCount = " + vHCount_In.ToString() + ", TicketIncome = " + vOAmt_In.ToString() + Environment.NewLine +
                                                               " where IndexNo = '" + vIndexNo_In + "' ";
                                            }
                                        }
                                        catch (Exception eMessage)
                                        {
                                            Response.Write("<Script language='Javascript'>");
                                            Response.Write("alert(" + eMessage.Message + ")");
                                            Response.Write("</" + "Script>");
                                        }
                                    }
                                }
                                else
                                {
                                    Response.Write("<Script language='Javascript'>");
                                    Response.Write("alert(" + vErrorStr + ")");
                                    Response.Write("</" + "Script>");
                                }
                            }
                        }
                        break;
                    case ".xlsx": //新版 EXCEL (2010 以後) 格式
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuTicketTrans.FileContent);
                        for (int si = 0; si < wbExcel_X.NumberOfSheets; si++)
                        {
                            XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(si);
                            if (sheetExcel_X.SheetName.Trim() == "清分明細報表")
                            {
                                //先進行一次資料檢查
                                for (int i = 0; i <= sheetExcel_X.LastRowNum; i++)
                                {
                                    //逐行讀入資料
                                    XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                                    if ((vRowTemp_X != null) && (vRowTemp_X.Cells.Count >= 16))
                                    {
                                        vSourceName = vRowTemp_X.Cells[0].ToString().Trim();
                                        if ((vSourceName != "") && (vSourceName != "票證公司"))
                                        {
                                            vBuDateStr_In = vRowTemp_X.Cells[1].ToString();
                                            if (DateTime.TryParse(vBuDateStr_In, out vGetDate_Temp))
                                            {
                                                vBuDateStr_In = ((vSourceName == "悠遊卡") || (vSourceName == "一卡通")) ? vGetDate_Temp.AddDays(1).ToString("yyyy/MM/dd") : vGetDate_Temp.ToString("yyyy/MM/dd");
                                            }
                                            else
                                            {
                                                vErrorStr += (Environment.NewLine + "'清分日期有誤，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            vDate_In = vRowTemp_X.Cells[2].ToString();
                                            if (DateTime.TryParse(vDate_In, out vGetDate_Temp))
                                            {
                                                vDate_In = vGetDate_Temp.ToString("yyyy/MM/dd");
                                            }
                                            else
                                            {
                                                vErrorStr += (Environment.NewLine + "'營運日期有誤，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            vDepNo_In = vRowTemp_X.Cells[3].ToString();
                                            vSQLStrTemp = "select isnull([Name], '') as [Name] from Department where DepNo = '" + vDepNo_In + "' ";
                                            if (PF.GetValue(vConnStr, vSQLStrTemp, "Name") == "")
                                            {
                                                vErrorStr += (Environment.NewLine + "'站別代碼有誤，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            vCarID_In = vRowTemp_X.Cells[4].ToString();
                                            if ((vCarID_In.Substring(0, 1) != "4") && (vCarID_In.IndexOf("-U7") > 0))
                                            {
                                                vCarID_In.Replace("-U7", "-U-7");
                                            }
                                            vSQLStrTemp = "select Car_No from Car_InfoA where Car_ID = '" + vCarID_In + "' ";
                                            if (PF.GetValue(vConnStr, vSQLStrTemp, "Car_No") == "")
                                            {
                                                vErrorStr += (Environment.NewLine + "'牌照號碼有誤，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            vDriver_In = Int32.Parse(vRowTemp_X.Cells[5].ToString()).ToString("D6");
                                            vSQLStrTemp = "select [Name] from Employee where EmpNo = '" + vDriver_In + "' ";
                                            if (PF.GetValue(vConnStr, vSQLStrTemp, "Name") == "")
                                            {
                                                vErrorStr += (Environment.NewLine + "'駕駛員工號有誤，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            vTicketLinesNo_In = vRowTemp_X.Cells[6].ToString();
                                            vSQLStrTemp = "select ERPLinesNo from LinesNoChart where TicketLineNo = '" + Int32.Parse(vTicketLinesNo_In).ToString() + "' ";
                                            if (PF.GetValue(vConnStr, vSQLStrTemp, "ERPLinesNo") == "")
                                            {
                                                vErrorStr += (Environment.NewLine + "'驗票機路線代碼有誤，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!double.TryParse(vRowTemp_X.Cells[8].ToString(), out vOAmt_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'營運金額必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            //2022/04/12 詢問主計課業管人員，目前票證公司提供的檔案中已經沒有聯營卡點數資料，聯營卡實際上也已經沒有使用了，所以固定帶 0
                                            //if (!double.TryParse(vRowTemp_X.Cells[8].ToString(), out vOPiece_In))
                                            //{
                                            //    vErrorStr += (Environment.NewLine + "'聯營卡點數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            //}
                                            if (!double.TryParse(vRowTemp_X.Cells[11].ToString(), out vSPiece_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'學生卡點數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!double.TryParse(vRowTemp_X.Cells[15].ToString(), out vOldPiece_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'敬老點數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!Int32.TryParse(vRowTemp_X.Cells[20].ToString(), out vHCount_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'總載客人數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!Int32.TryParse(vRowTemp_X.Cells[21].ToString(), out vPCount_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'普卡人數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!Int32.TryParse(vRowTemp_X.Cells[22].ToString(), out vLCount_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'聯營卡人數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!Int32.TryParse(vRowTemp_X.Cells[25].ToString(), out vSCount_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'學生卡人數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                            if (!Int32.TryParse(vRowTemp_X.Cells[28].ToString(), out vOldCount_In))
                                            {
                                                vErrorStr += (Environment.NewLine + "'敬老人數必須是數字，資料筆數：" + (i + 1).ToString() + "' ");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        vErrorStr += (Environment.NewLine + "'EXCEL 檔案格式有誤!' ");
                                    }
                                }
                                if (vErrorStr.Trim() == "")
                                {
                                    for (int i = 0; i <= sheetExcel_X.LastRowNum; i++)
                                    {
                                        //逐行讀入資料
                                        XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                                        vSourceName = vRowTemp_X.Cells[0].ToString().Trim();
                                        if ((vSourceName != "") && (vSourceName != "票證公司"))
                                        {
                                            vBuDateStr_In = vRowTemp_X.Cells[1].ToString();
                                            vGetDate_Temp = DateTime.Parse(vBuDateStr_In);
                                            vBuDateStr_In = ((vSourceName == "悠遊卡") || (vSourceName == "一卡通")) ? vGetDate_Temp.AddDays(1).ToString("yyyy/MM/dd") : vGetDate_Temp.ToString("yyyy/MM/dd");
                                            vCalYM_In = (DateTime.Parse(vBuDateStr_In).Year - 1911).ToString("D3") + (DateTime.Parse(vBuDateStr_In).Month.ToString("D2"));
                                            vDate_In = vRowTemp_X.Cells[2].ToString();
                                            vGetDate_Temp = DateTime.Parse(vDate_In);
                                            vDate_In = vGetDate_Temp.ToString("yyyy/MM/dd");
                                            vDepNo_In = vRowTemp_X.Cells[3].ToString();
                                            vCarID_In = vRowTemp_X.Cells[4].ToString();
                                            if ((vCarID_In.Substring(0, 1) != "4") && (vCarID_In.IndexOf("-U7") > 0))
                                            {
                                                vCarID_In.Replace("-U7", "-U-7");
                                            }
                                            vSQLStrTemp = "select Car_No from Car_InfoA where Car_ID = '" + vCarID_In + "' ";
                                            vCarNo_In = PF.GetValue(vConnStr, vSQLStrTemp, "Car_No");
                                            vDriver_In = Int32.Parse(vRowTemp_X.Cells[5].ToString()).ToString("D6");
                                            vTicketLinesNo_In = vRowTemp_X.Cells[6].ToString();
                                            vSQLStrTemp = "select ERPLinesNo from LinesNoChart where TicketLineNo = '" + Int32.Parse(vTicketLinesNo_In).ToString() + "' ";
                                            vLinesNo_In = PF.GetValue(vConnStr, vSQLStrTemp, "ERPLinesNo");
                                            vIndexNo_In = vCalYM_In + vTicketType_In + vTicketLinesNo_In;
                                            vOAmt_In = double.Parse(vRowTemp_X.Cells[8].ToString());
                                            //2022/04/12 詢問主計課業管人員，目前票證公司提供的檔案中已經沒有聯營卡點數資料，聯營卡實際上也已經沒有使用了，所以固定帶 0
                                            vOPiece_In = 0.0;
                                            vSPiece_In = double.Parse(vRowTemp_X.Cells[11].ToString());
                                            vOldPiece_In = double.Parse(vRowTemp_X.Cells[15].ToString());
                                            vHCount_In = Int32.Parse(vRowTemp_X.Cells[20].ToString());
                                            vPCount_In = Int32.Parse(vRowTemp_X.Cells[21].ToString());
                                            vLCount_In = Int32.Parse(vRowTemp_X.Cells[22].ToString());
                                            vSCount_In = Int32.Parse(vRowTemp_X.Cells[25].ToString());
                                            vOldCount_In = Int32.Parse(vRowTemp_X.Cells[28].ToString());
                                        }
                                        try
                                        {
                                            vSQLStr_Main = "insert into CPSOutG " + Environment.NewLine +
                                                           "            (BuDate, Date, DepNo, LinesNo, Car_ID, Car_No, Driver, OAmt, OPiece, SPiece, " + Environment.NewLine +
                                                           "             HCount, SCount, OldPiece, PCount, LCount, OldCount)" + Environment.NewLine +
                                                           "      values(@BuDate, @Date, @DepNo, @LinesNo, @Car_ID, @Car_No, @Driver, @OAmt, @OPiece, @SPiece, " + Environment.NewLine +
                                                           "             @HCount, @SCount, @OldPiece, @PCount, @LCount, @OldCount)";
                                            sdsCPSOutG_Detail.InsertCommand = vSQLStr_Main;
                                            sdsCPSOutG_Detail.InsertParameters.Clear();
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, vBuDateStr_In));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("Date", DbType.Date, vDate_In));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo_In));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("Car_ID", DbType.String, vCarID_In));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("Car_No", DbType.String, vCarNo_In));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("Driver", DbType.String, vDriver_In));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("OAmt", DbType.Double, vOAmt_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("OPiece", DbType.Double, vOPiece_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("SPiece", DbType.Double, vSPiece_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("OldPiece", DbType.Double, vOldPiece_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("HCount", DbType.Double, vHCount_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("SCount", DbType.Double, vSCount_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("PCount", DbType.Double, vPCount_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("LCount", DbType.Double, vLCount_In.ToString()));
                                            sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("OldCount", DbType.Double, vOldCount_In.ToString()));
                                            sdsCPSOutG_Detail.Insert();
                                            gridCPSOutG.DataBind();

                                            vSQLStrTemp = "select IndexNo from LinesBounds where IndexNo = '" + vIndexNo_In + "' ";
                                            if (PF.GetValue(vConnStr, vSQLStrTemp, "IndeexNo") == "")
                                            {
                                                vSQLStr_Main = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLinesNo, LinesGOVNo, TicketPCount, TicketIncome)" + Environment.NewLine +
                                                               "     values ('" + vIndexNo_In + "', '" + vCalYM_In + "', '" + vTicketType_In + "', '" + vTicketLinesNo_In + "', '" + vGovLinesNo_In + "', " + vHCount_In.ToString() + ", " + vOAmt_In.ToString() + ")";
                                            }
                                            else
                                            {
                                                vSQLStr_Main = "update LinesBounds set TicketPCount = " + vHCount_In.ToString() + ", TicketIncome = " + vOAmt_In.ToString() + Environment.NewLine +
                                                               " where IndexNo = '" + vIndexNo_In + "' ";
                                            }
                                        }
                                        catch (Exception eMessage)
                                        {
                                            Response.Write("<Script language='Javascript'>");
                                            Response.Write("alert(" + eMessage.Message + ")");
                                            Response.Write("</" + "Script>");
                                        }
                                    }
                                }
                                else
                                {
                                    Response.Write("<Script language='Javascript'>");
                                    Response.Write("alert(" + vErrorStr + ")");
                                    Response.Write("</" + "Script>");
                                }
                            }
                        }
                        break;
                    case ".csv": //用逗號分隔的 CSV 檔
                        var vSourceFile = new StreamReader(fuTicketTrans.FileContent);
                        //先跑一次檢查資料內容有沒有問題
                        while (!vSourceFile.EndOfStream)
                        {
                            vIndex++;
                            string vSourceData = vSourceFile.ReadLine();
                            if (vSourceData.Trim() != "")
                            {
                                string[] vaSourceData = vSourceData.Split(',');
                                vBuDateStr_In = vaSourceData[0].Trim();
                                if (DateTime.TryParse(vBuDateStr_In, out vGetDate_Temp))
                                {
                                    vBuDateStr_In = vGetDate_Temp.ToString("yyyy/MM/dd");
                                }
                                else
                                {
                                    vErrorStr += (Environment.NewLine + "'清分日期有誤，資料筆數：" + vIndex.ToString() + "' ");
                                }
                                vDate_In = vaSourceData[1].Trim();
                                if (DateTime.TryParse(vDate_In, out vGetDate_Temp))
                                {
                                    vDate_In = vGetDate_Temp.ToString("yyyy/MM/dd");
                                }
                                else
                                {
                                    vErrorStr += (Environment.NewLine + "'營運日期有誤，資料筆數：" + vIndex.ToString() + "' ");
                                }
                                vDepNo_In = vaSourceData[2].Trim();
                                vSQLStrTemp = "select isnull([Name], '') as [Name] from Department where DepNo = '" + vDepNo_In + "' ";
                                if (PF.GetValue(vConnStr, vSQLStrTemp, "Name") == "")
                                {
                                    vErrorStr += (Environment.NewLine + "'站別代碼有誤，資料筆數：" + vIndex.ToString() + "' ");
                                }
                                vTicketLinesNo_In = vaSourceData[3].Trim();
                                //驗票機路線代號查驗方式後補
                                vCarID_In = vaSourceData[4].Trim();
                                if ((vCarID_In.Substring(0, 1) != "4") && (vCarID_In.IndexOf("-U7") > 0))
                                {
                                    vCarID_In.Replace("-U7", "-U-7");
                                }
                                vSQLStrTemp = "select Car_No from Car_InfoA where Car_ID = '" + vCarID_In + "' ";
                                if (PF.GetValue(vConnStr, vSQLStrTemp, "Car_No") == "")
                                {
                                    vErrorStr += (Environment.NewLine + "'牌照號碼有誤，資料筆數：" + vIndex.ToString() + "' ");
                                }
                                vDriver_In = Int32.Parse(vaSourceData[5].Trim()).ToString("D2") + Int32.Parse(vaSourceData[6].Trim()).ToString("D4");
                                vSQLStrTemp = "select [Name] from Employee where EmpNo = '" + vDriver_In + "' ";
                                if (PF.GetValue(vConnStr, vSQLStrTemp, "Name") == "")
                                {
                                    vErrorStr += (Environment.NewLine + "'駕駛員工號有誤，資料筆數：" + vIndex.ToString() + "' ");
                                }
                                if (!double.TryParse(vaSourceData[7].Trim(), out vOAmt_In))
                                {
                                    vErrorStr += (Environment.NewLine + "'營運金額必須是數字，資料筆數：" + vIndex.ToString() + "' ");
                                }
                                if (!double.TryParse(vaSourceData[8].Trim(), out vOPiece_In))
                                {
                                    vErrorStr += (Environment.NewLine + "'聯營卡點數必須是數字，資料筆數：" + vIndex.ToString() + "' ");
                                }
                                if (!double.TryParse(vaSourceData[9].Trim(), out vSPiece_In))
                                {
                                    vErrorStr += (Environment.NewLine + "'學生卡點數必須是數字，資料筆數：" + vIndex.ToString() + "' ");
                                }
                                if (!double.TryParse(vaSourceData[10].Trim(), out vOldPiece_In))
                                {
                                    vErrorStr += (Environment.NewLine + "'敬老點數必須是數字，資料筆數：" + vIndex.ToString() + "' ");
                                }
                                if (!Int32.TryParse(vaSourceData[11].Trim(), out vHCount_In))
                                {
                                    vErrorStr += (Environment.NewLine + "'總載客人數必須是數字，資料筆數：" + vIndex.ToString() + "' ");
                                }
                                if (!Int32.TryParse(vaSourceData[12].Trim(), out vPCount_In))
                                {
                                    vErrorStr += (Environment.NewLine + "'普卡人數必須是數字，資料筆數：" + vIndex.ToString() + "' ");
                                }
                                if (!Int32.TryParse(vaSourceData[13].Trim(), out vLCount_In))
                                {
                                    vErrorStr += (Environment.NewLine + "'聯營卡人數必須是數字，資料筆數：" + vIndex.ToString() + "' ");
                                }
                                if (!Int32.TryParse(vaSourceData[14].Trim(), out vSCount_In))
                                {
                                    vErrorStr += (Environment.NewLine + "'學生卡人數必須是數字，資料筆數：" + vIndex.ToString() + "' ");
                                }
                                if (!Int32.TryParse(vaSourceData[15].Trim(), out vOldCount_In))
                                {
                                    vErrorStr += (Environment.NewLine + "'敬老人數必須是數字，資料筆數：" + vIndex.ToString() + "' ");
                                }
                            }
                        }
                        if (vErrorStr.Trim() == "")
                        {
                            //先關閉 StreamReader 再重新讀取一次
                            vSourceFile.Close();
                            vSourceFile = new StreamReader(fuTicketTrans.FileContent);
                            while (!vSourceFile.EndOfStream)
                            {
                                string vSourceData = vSourceFile.ReadLine();
                                if (vSourceData.Trim() != "")
                                {
                                    string[] vaSourceData = vSourceData.Split(',');
                                    vBuDateStr_In = vaSourceData[0].Trim();
                                    vDate_In = vaSourceData[1].Trim();
                                    vDepNo_In = vaSourceData[2].Trim();
                                    vTicketLinesNo_In = vaSourceData[3].Trim();
                                    vCarID_In = vaSourceData[4].Trim();
                                    if ((vCarID_In.Substring(0, 1) != "4") && (vCarID_In.IndexOf("-U7") > 0))
                                    {
                                        vCarID_In.Replace("-U7", "-U-7");
                                    }
                                    vSQLStrTemp = "select Car_No from Car_InfoA where Car_ID = '" + vCarID_In + "' ";
                                    vCarNo_In = PF.GetValue(vConnStr, vSQLStrTemp, "Car_No");
                                    vDriver_In = Int32.Parse(vaSourceData[5].Trim()).ToString("D2") + Int32.Parse(vaSourceData[6].Trim()).ToString("D4");
                                    if (!double.TryParse(vaSourceData[7].Trim(), out vOAmt_In))
                                    {
                                        vOAmt_In = 0.0;
                                    }
                                    if (!double.TryParse(vaSourceData[8].Trim(), out vOPiece_In))
                                    {
                                        vOPiece_In = 0.0;
                                    }
                                    if (!double.TryParse(vaSourceData[9].Trim(), out vSPiece_In))
                                    {
                                        vSPiece_In = 0.0;
                                    }
                                    if (!double.TryParse(vaSourceData[10].Trim(), out vOldPiece_In))
                                    {
                                        vOldPiece_In = 0.0;
                                    }
                                    if (!Int32.TryParse(vaSourceData[11].Trim(), out vHCount_In))
                                    {
                                        vHCount_In = 0;
                                    }
                                    if (!Int32.TryParse(vaSourceData[12].Trim(), out vPCount_In))
                                    {
                                        vPCount_In = 0;
                                    }
                                    if (!Int32.TryParse(vaSourceData[13].Trim(), out vLCount_In))
                                    {
                                        vLCount_In = 0;
                                    }
                                    if (!Int32.TryParse(vaSourceData[14].Trim(), out vSCount_In))
                                    {
                                        vSCount_In = 0;
                                    }
                                    if (!Int32.TryParse(vaSourceData[15].Trim(), out vOldCount_In))
                                    {
                                        vOldCount_In = 0;
                                    }
                                }
                            }
                            try
                            {
                                vSQLStr_Main = "insert into CPSOutG " + Environment.NewLine +
                                               "            (BuDate, Date, DepNo, LinesNo, Car_ID, Car_No, Driver, OAmt, OPiece, SPiece, " + Environment.NewLine +
                                               "             HCount, SCount, OldPiece, PCount, LCount, OldCount)" + Environment.NewLine +
                                               "      values(@BuDate, @Date, @DepNo, @LinesNo, @Car_ID, @Car_No, @Driver, @OAmt, @OPiece, @SPiece, " + Environment.NewLine +
                                               "             @HCount, @SCount, @OldPiece, @PCount, @LCount, @OldCount)";
                                sdsCPSOutG_Detail.InsertCommand = vSQLStr_Main;
                                sdsCPSOutG_Detail.InsertParameters.Clear();
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, vBuDateStr_In));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("Date", DbType.Date, vDate_In));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo_In));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("Car_ID", DbType.String, vCarID_In));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("Car_No", DbType.String, vCarNo_In));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("Driver", DbType.String, vDriver_In));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("OAmt", DbType.Double, vOAmt_In.ToString()));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("OPiece", DbType.Double, vOPiece_In.ToString()));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("SPiece", DbType.Double, vSPiece_In.ToString()));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("OldPiece", DbType.Double, vOldPiece_In.ToString()));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("HCount", DbType.Double, vHCount_In.ToString()));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("SCount", DbType.Double, vSCount_In.ToString()));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("PCount", DbType.Double, vPCount_In.ToString()));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("LCount", DbType.Double, vLCount_In.ToString()));
                                sdsCPSOutG_Detail.InsertParameters.Add(new Parameter("OldCount", DbType.Double, vOldCount_In.ToString()));
                                sdsCPSOutG_Detail.Insert();
                                gridCPSOutG.DataBind();
                            }
                            catch (Exception eMessage)
                            {
                                Response.Write("<Script language='Javascript'>");
                                Response.Write("alert(" + eMessage.Message + ")");
                                Response.Write("</" + "Script>");
                            }
                        }
                        else
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert(" + vErrorStr + ")");
                            Response.Write("</" + "Script>");
                        }
                        break;
                    default:
                        break;
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇資料的來源檔案')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 刪除清分資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbClear_Search_Click(object sender, EventArgs e)
        {
            string vDelStr = "";
            if ((eBuDate_S_Search.Text.Trim() == "") && (eBuDate_E_Search.Text.Trim() == ""))
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先輸入要刪除清分資料的日期區間')");
                Response.Write("</" + "Script>");
            }
            else if ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() != ""))
            {
                vDelStr = "delete CPSOutG where BuDate between @BuDate_S and '@BuDate_E ";
                sdsCPSOutG_Detail.DeleteCommand = vDelStr;
                sdsCPSOutG_Detail.DeleteParameters.Clear();
                sdsCPSOutG_Detail.DeleteParameters.Add(new Parameter("BuDate_S", DbType.Date, eBuDate_S_Search.Text.Trim()));
                sdsCPSOutG_Detail.DeleteParameters.Add(new Parameter("BuDate_E", DbType.Date, eBuDate_E_Search.Text.Trim()));
                sdsCPSOutG_Detail.Delete();
            }
            else if ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() == ""))
            {
                vDelStr = "delete CPSOutG where BuDate = @Budate";
                sdsCPSOutG_Detail.DeleteCommand = vDelStr;
                sdsCPSOutG_Detail.DeleteParameters.Clear();
                sdsCPSOutG_Detail.DeleteParameters.Add(new Parameter("BuDate", DbType.Date, eBuDate_S_Search.Text.Trim()));
                sdsCPSOutG_Detail.Delete();
            }
            else if ((eBuDate_S_Search.Text.Trim() == "") && (eBuDate_E_Search.Text.Trim() != ""))
            {
                vDelStr = "delete CPSOutG where BuDate = @Budate";
                sdsCPSOutG_Detail.DeleteCommand = vDelStr;
                sdsCPSOutG_Detail.DeleteParameters.Clear();
                sdsCPSOutG_Detail.DeleteParameters.Add(new Parameter("BuDate", DbType.Date, eBuDate_E_Search.Text.Trim()));
                sdsCPSOutG_Detail.Delete();
            }
            gridCPSOutG.DataBind();
            fvCPSOutG.DataBind();
        }

        protected void bbClose_Search_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}