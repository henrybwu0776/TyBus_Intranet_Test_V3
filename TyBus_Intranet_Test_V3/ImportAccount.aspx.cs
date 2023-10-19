using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class ImportAccount : System.Web.UI.Page
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
            string vSQLStr = "";
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
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        eAccDate_Search.Text = DateTime.Today.ToShortDateString();
                        eIgoneExcelDate_Search.Checked = true;
                        /*
                        vSQLStr = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                  " union all " + Environment.NewLine +
                                  "select ClassNo, ClassTxt from DBDICB where FKey = '傳票日記帳      ACCOUNT         TYPE' order by ClassNo"; //*/
                        vSQLStr = "select ClassNo, ClassTxt from DBDICB where FKey = '傳票日記帳      ACCOUNT         TYPE' order by ClassNo";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                ddlAccType_Search.Items.Add(new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim()));
                            }
                        }
                        ddlAccType_Search.SelectedIndex = 0;
                        plShowData.Visible = false;
                    }
                    else
                    {

                    }
                    string vAccDateURL = "InputDate.aspx?TextBoxID=" + eAccDate_Search.ClientID;
                    string vAccDateScript = "window.open('" + vAccDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eAccDate_Search.Attributes["onClick"] = vAccDateScript;
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

        private DataTable CreateTempDataTable()
        {
            DataTable dtTemp = new DataTable();
            //唯一值
            DataColumn dcSID = new DataColumn("SID", typeof(String));
            dcSID.MaxLength = 40; //資料長度
            dtTemp.Columns.Add(dcSID);
            //傳票單號
            DataColumn dcAccountNo = new DataColumn("No", typeof(String));
            dcAccountNo.MaxLength = 15;
            //dcAccountNo.Unique = false; //是否可以重複，預設 false 表示可以重複
            //dcAccountNo.AllowDBNull = true; //是否可以為空值，預設 true 表示可以為空值
            dtTemp.Columns.Add(dcAccountNo);
            //傳票種類
            DataColumn dcType = new DataColumn("Type", typeof(String));
            dcType.MaxLength = 1;
            dtTemp.Columns.Add(dcType);
            //建檔日期
            DataColumn dcHap_Date = new DataColumn("Hap_Date", typeof(DateTime));
            dtTemp.Columns.Add(dcHap_Date);
            //會計人員
            DataColumn dcAcc_Man = new DataColumn("Acc_Man", typeof(String));
            dcAcc_Man.MaxLength = 12;
            dcAcc_Man.AllowDBNull = false;
            dtTemp.Columns.Add(dcAcc_Man);
            //列印次數
            DataColumn dcPrint_CNT = new DataColumn("Print_CNT", typeof(Double));
            dcPrint_CNT.AllowDBNull = false;
            dtTemp.Columns.Add(dcPrint_CNT);
            //傳票日期
            DataColumn dcRec_Date = new DataColumn("Rec_Date", typeof(DateTime));
            dtTemp.Columns.Add(dcRec_Date);
            //明細項次
            DataColumn dcItems = new DataColumn("Item", typeof(String));
            dcItems.MaxLength = 5;
            dcItems.AllowDBNull = false;
            dtTemp.Columns.Add(dcItems);
            //會計科目
            DataColumn dcSubject = new DataColumn("Subject", typeof(String));
            dcSubject.MaxLength = 12;
            dcSubject.AllowDBNull = false;
            dtTemp.Columns.Add(dcSubject);
            //客戶廠商
            DataColumn dcTarget = new DataColumn("Target", typeof(String));
            dcTarget.MaxLength = 12;
            dtTemp.Columns.Add(dcTarget);
            //
            DataColumn dcStk_Month = new DataColumn("Stk_Month", typeof(DateTime));
            dtTemp.Columns.Add(dcStk_Month);
            //部門編號
            DataColumn dcDepNo = new DataColumn("Department", typeof(String));
            dcDepNo.MaxLength = 12;
            dtTemp.Columns.Add(dcDepNo);
            //摘要
            DataColumn dcMemo = new DataColumn("Meno", typeof(String));
            dcMemo.MaxLength = 80;
            dtTemp.Columns.Add(dcMemo);
            //借方金額
            DataColumn dcDebit = new DataColumn("Debit", typeof(Double));
            dtTemp.Columns.Add(dcDebit);
            //貸方金額
            DataColumn dcCredit = new DataColumn("Credit", typeof(Double));
            dtTemp.Columns.Add(dcCredit);
            //
            DataColumn dcStk_Type = new DataColumn("Stk_Type", typeof(String));
            dcStk_Type.MaxLength = 1;
            dtTemp.Columns.Add(dcStk_Type);
            //
            DataColumn dcStk_Date1 = new DataColumn("Stk_Date1", typeof(DateTime));
            dtTemp.Columns.Add(dcStk_Date1);
            //
            DataColumn dcStk_No = new DataColumn("Stk_No", typeof(String));
            dcStk_No.MaxLength = 15;
            dtTemp.Columns.Add(dcStk_No);
            //
            DataColumn dcStk_Item = new DataColumn("Stk_Item", typeof(String));
            dcStk_Item.MaxLength = 5;
            dtTemp.Columns.Add(dcStk_Item);
            //
            DataColumn dcBStk_Date = new DataColumn("BStk_Date", typeof(DateTime));
            dtTemp.Columns.Add(dcBStk_Date);
            //
            DataColumn dcBStk_No = new DataColumn("BStk_No", typeof(String));
            dcBStk_No.MaxLength = 15;
            dtTemp.Columns.Add(dcBStk_No);
            //
            DataColumn dcBStk_Item = new DataColumn("BStk_Item", typeof(String));
            dcBStk_Item.MaxLength = 5;
            dtTemp.Columns.Add(dcBStk_Item);
            //發票號碼
            DataColumn dcInvoice = new DataColumn("Invoice", typeof(String));
            dcInvoice.MaxLength = 10;
            dtTemp.Columns.Add(dcInvoice);
            //
            DataColumn dcEnd_Date = new DataColumn("End_Date", typeof(DateTime));
            dtTemp.Columns.Add(dcEnd_Date);
            //立沖金額
            DataColumn dcStk_Money = new DataColumn("Stk_Money", typeof(Double));
            dtTemp.Columns.Add(dcStk_Money);
            //
            DataColumn dcMod_Type = new DataColumn("Mod_Type", typeof(String));
            dcMod_Type.MaxLength = 1;
            dtTemp.Columns.Add(dcMod_Type);
            //
            DataColumn dcMod_Class = new DataColumn("Mod_Class", typeof(String));
            dcMod_Class.MaxLength = 1;
            dtTemp.Columns.Add(dcMod_Class);
            //
            DataColumn dcMod_Man = new DataColumn("Mod_Man", typeof(String));
            dcMod_Man.MaxLength = 8;
            dtTemp.Columns.Add(dcMod_Man);
            //
            DataColumn dcMod_Date = new DataColumn("Mod_Date", typeof(DateTime));
            dtTemp.Columns.Add(dcMod_Date);
            //
            DataColumn dcProduct = new DataColumn("Product", typeof(String));
            dcProduct.MaxLength = 20;
            dtTemp.Columns.Add(dcProduct);
            //匯率
            DataColumn dcRate = new DataColumn("Rate", typeof(Double));
            dtTemp.Columns.Add(dcRate);
            //原幣借方
            DataColumn dcFor_Debit = new DataColumn("For_Debit", typeof(Double));
            dtTemp.Columns.Add(dcFor_Debit);
            //原幣貸方
            DataColumn dcFor_Credit = new DataColumn("For_Credit", typeof(Double));
            dtTemp.Columns.Add(dcFor_Credit);
            //原幣幣別
            DataColumn dcFor_Type = new DataColumn("For_Type", typeof(String));
            dcFor_Type.MaxLength = 4;
            dtTemp.Columns.Add(dcFor_Type);
            //來源類別
            DataColumn dcSource = new DataColumn("Source", typeof(String));
            dcSource.MaxLength = 10;
            dtTemp.Columns.Add(dcSource);
            //來源單號
            DataColumn dcSourceID = new DataColumn("SourceID", typeof(String));
            dcSourceID.MaxLength = 16;
            dtTemp.Columns.Add(dcSourceID);
            //
            DataColumn dcBlock = new DataColumn("Block", typeof(String));
            dcBlock.MaxLength = 1;
            dtTemp.Columns.Add(dcBlock);
            //
            DataColumn dcAccType = new DataColumn("AccType", typeof(String));
            dcAccType.MaxLength = 4;
            dtTemp.Columns.Add(dcAccType);
            //
            DataColumn dcCFUser = new DataColumn("CFUser", typeof(String));
            dcCFUser.MaxLength = 12;
            dtTemp.Columns.Add(dcCFUser);
            //
            DataColumn dcCFDate = new DataColumn("CFDate", typeof(DateTime));
            dtTemp.Columns.Add(dcCFDate);
            //
            DataColumn dcVoidMan = new DataColumn("VoidMan", typeof(String));
            dcVoidMan.MaxLength = 12;
            dtTemp.Columns.Add(dcVoidMan);
            //
            DataColumn dcVoidDate = new DataColumn("VoidDate", typeof(DateTime));
            dtTemp.Columns.Add(dcVoidDate);
            //
            DataColumn dcCarID = new DataColumn("Car_ID", typeof(String));
            dcCarID.MaxLength = 12;
            dtTemp.Columns.Add(dcCarID);
            //
            DataColumn dcCarNo = new DataColumn("Car_No", typeof(String));
            dcCarNo.MaxLength = 12;
            dtTemp.Columns.Add(dcCarNo);
            //摘要2
            DataColumn dcMemo2 = new DataColumn("Memo2", typeof(String));
            dcMemo2.MaxLength = 20;
            dtTemp.Columns.Add(dcMemo2);
            //票據張數
            DataColumn dcDocCount = new DataColumn("DocCount", typeof(Int32));
            dtTemp.Columns.Add(dcDocCount);

            return dtTemp;
        }

        private void ImportData_H()
        {
            //Excel 97-2003
            string vSQLStr_Temp = "";
            DateTime vAcc_Date_Temp = DateTime.Parse(eAccDate_Search.Text.Trim());
            DateTime vAcc_Date_Source;
            DateTime vAcc_Date_Close;
            DateTime vTempDate;
            string vAccountNo = "";
            string vAccountYM = "";
            string vNoString = "";
            string vSubject = "";
            int vDebit = 0;
            int vCredit = 0;
            int vTotalDebit = 0;
            int vTotalCredit = 0;
            string vDepNo = "";
            string vDepName = "";
            string vRemark = "";
            string vMemo2 = "";
            string vStoreNo = "";
            string vAccDateStr = "";
            int vSheetCount = 0;
            string vClass1 = "";
            int vItemMax = 0;
            int vItems = 0;
            string vNoMax = "";
            string vAcc_Source = "";
            string vAcc_SourceID = "";
            string vTypeStr = "";
            string vTempStr = "";
            int vSubject1110 = 0;

            //取回關帳日期
            vSQLStr_Temp = "select Desscription from AC_CLS where PK = '關帳' and Class = '關帳日期' ";
            vAcc_Date_Close = DateTime.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "Desscription"));

            HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
            vSheetCount = wbExcel_H.NumberOfSheets;
            dtTarget = CreateTempDataTable();
            HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
            HSSFRow vFirstRow = (HSSFRow)sheetExcel_H.GetRow(sheetExcel_H.FirstRowNum);
            vNoString = vFirstRow.GetCell(0).ToString().Trim();
            switch (ddlAccType_Search.SelectedValue)
            {
                default:
                    vTypeStr = "0";
                    break;

                case "2":
                    vTypeStr = "4";
                    break;

                case "3":
                    vTypeStr = "5";
                    break;
            }
            for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
            {
                HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                vTempStr = "";
                if (vRowTemp_H.GetCell(0).ToString().Trim() != vNoString)
                {
                    //NoString 有變動，表示換了一張傳票，所以要把前一張傳票的資料同步進資料庫
                    //先檢查傳票借貸是否平衡
                    if ((vTotalDebit != vTotalCredit) && (dtTarget.Rows.Count > 0))
                    {
                        vErrorMessage += "EXCEL 傳票：[" + vNoString.Trim() + "] 明細借貸不平！" + Environment.NewLine;
                    }
                    //檢查現收或現支傳票是不是有至少一筆現金科目
                    if (((ddlAccType_Search.SelectedValue != "1") && (vSubject1110 == 0)) && (dtTarget.Rows.Count > 0))
                    {
                        vErrorMessage += ddlAccType_Search.SelectedItem.Text.Trim() + "傳票缺少必要的現金科目！" + Environment.NewLine;
                    }
                    //確定這張傳票沒有錯誤訊息時才做同步的動作
                    if (((vErrorMessage == null) || (vErrorMessage.Length == 0)) && (dtTarget.Rows.Count > 0))
                    {
                        //先把日期欄位的字串轉成沒有分隔符號的民國年月日字串
                        vTempDate = DateTime.Parse(vAccDateStr);
                        vAccountYM = (vTempDate.Year - 1911).ToString("D3") + vTempDate.Month.ToString("D2") + vTempDate.Day.ToString("D2") + vTypeStr;
                        //取回同一個日期字串的最後一張傳票的號碼
                        vSQLStr_Temp = "select Max([No]) MaxNo from Account where [No] like '" + vAccountYM + "%' ";
                        vNoMax = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                        //計算新的傳票號碼
                        if (Int32.TryParse(vNoMax.Replace(vAccountYM, ""), out vItemMax))
                        {
                            vItemMax++;
                        }
                        else
                        {
                            vItemMax = 1;
                        }
                        vAccountNo = vAccountYM + vItemMax.ToString("D3");
                        for (int TableRowCount = 0; TableRowCount < dtTarget.Rows.Count; TableRowCount++)
                        {
                            dtTarget.Rows[TableRowCount]["No"] = vAccountNo;
                        }
                        using (SqlBulkCopy sbkTemp = new SqlBulkCopy(vConnStr, SqlBulkCopyOptions.UseInternalTransaction))
                        {
                            try
                            {
                                sbkTemp.DestinationTableName = "Account";
                                sbkTemp.WriteToServer(dtTarget);
                                dtTarget.Clear();
                            }
                            catch (Exception eMessage)
                            {
                                Response.Write("<Script language='Javascript'>");
                                Response.Write("alert('" + eMessage.Message + "')");
                                Response.Write("</" + "Script>");
                            }
                        }
                    }
                    //同步完之後更新 NoString 的值
                    vNoString = vRowTemp_H.GetCell(0).ToString().Trim();
                    vItems = 0;
                    vTotalCredit = 0;
                    vTotalDebit = 0;
                    vSubject1110 = 0;
                }
                if (vNoString != "傳票單號")
                //if (vRowTemp_H.GetCell(0).ToString().Trim() != "傳票單號")
                {
                    //2023.10.02 因應測試回饋修改
                    //vAccDateStr = (eIgoneExcelDate_Search.Checked == true) ? vAcc_Date_Temp.ToShortDateString() : vRowTemp_H.GetCell(16).ToString().Trim();
                    vTempDate = (eIgoneExcelDate_Search.Checked == true) ? vAcc_Date_Temp : DateTime.Parse(vRowTemp_H.GetCell(16).ToString().Trim());
                    vAccDateStr = vTempDate.ToString("yyyy/MM/dd") + " 00:00:00";
                    vSubject = vRowTemp_H.GetCell(2).ToString().Trim();
                    vDebit = (int)vRowTemp_H.GetCell(4).NumericCellValue;
                    vCredit = (int)vRowTemp_H.GetCell(5).NumericCellValue;
                    vDepNo = vRowTemp_H.GetCell(6).ToString().Trim();
                    vRemark = vRowTemp_H.GetCell(8).ToString().Trim();
                    vMemo2 = vRowTemp_H.GetCell(9).ToString().Trim();
                    vStoreNo = vRowTemp_H.GetCell(10).ToString().Trim();
                    vAcc_Source = (vRowTemp_H.GetCell(22).ToString().Trim() != "") ? vRowTemp_H.GetCell(22).ToString().Trim() : String.Empty;
                    vAcc_SourceID = (vRowTemp_H.GetCell(23).ToString().Trim() != "") ? vRowTemp_H.GetCell(23).ToString().Trim() : String.Empty;

                    //檢查傳票日期
                    if (!DateTime.TryParse(vAccDateStr, out vAcc_Date_Source))
                    {
                        vErrorMessage += "EXCEL 行數：[" + i.ToString().Trim() + "] 傳票日期有誤！" + Environment.NewLine;
                    }
                    else if (vAcc_Date_Source < vAcc_Date_Close)
                    {
                        vErrorMessage += "EXCEL 行數：[" + i.ToString().Trim() + "] 指定日期已關帳！" + Environment.NewLine;
                    }

                    //取回是否是統制科目
                    vSQLStr_Temp = "select Class1 from AC_Subject where Subject = '" + vSubject + "' ";
                    vClass1 = PF.GetValue(vConnStr, vSQLStr_Temp, "Class1");
                    if (vClass1.Trim() == "")
                    {
                        vErrorMessage += "EXCEL 行數：[" + i.ToString().Trim() + "] 找不到對應的會計科目！" + Environment.NewLine;
                    }
                    else if (vClass1 == "1")
                    {
                        vErrorMessage += "EXCEL 行數：[" + i.ToString().Trim() + "] 會計科目為統制科目！" + Environment.NewLine;
                    }
                    //檢查站別編號
                    if (vDepNo != "")
                    {
                        vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                        vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                        if (vDepName == "")
                        {
                            vErrorMessage += "EXCEL 行數：[" + i.ToString().Trim() + "] 站別有誤！" + Environment.NewLine;
                        }
                    }
                    //檢查現金科目 (111001) 內容
                    if (vSubject == "111001")
                    {
                        switch (ddlAccType_Search.SelectedValue)
                        {
                            default:
                                vTempStr = "";
                                break;

                            case "2":
                                vTempStr = (vDebit < 0) ? "EXCEL 行數：[" + i.ToString().Trim() + "] 現金收入傳票中現金科目借方金額有誤！" : "";
                                vSubject1110++;
                                break;

                            case "3":
                                vTempStr = (vCredit < 0) ? "EXCEL 行數：[" + i.ToString().Trim() + "] 現金支出傳票中現金科目貸方金額有誤！" : "";
                                vSubject1110++;
                                break;
                        }
                        if (vTempStr.Trim() != "")
                        {
                            vErrorMessage += vTempStr + Environment.NewLine;
                        }
                    }

                    /*
                    //檢查金額不可為負數
                    if ((vDebit < 0) || (vCredit < 0))
                    {
                        vaErrorMessage += "EXCEL 行數：[" + i.ToString().Trim() + "] 借方或貸方金額不可為負數！" + Environment.NewLine;
                    } //*/

                    vItems++;
                    vTotalDebit += vDebit;
                    vTotalCredit += vCredit;
                    DataRow rowTemp = dtTarget.NewRow();
                    rowTemp["Item"] = vItems.ToString("D4");
                    rowTemp["Subject"] = vSubject;
                    rowTemp["Debit"] = vDebit;
                    rowTemp["Credit"] = vCredit;
                    rowTemp["Department"] = vDepNo;
                    rowTemp["Meno"] = vRemark;
                    rowTemp["Memo2"] = vMemo2;
                    rowTemp["Target"] = vStoreNo;
                    rowTemp["Type"] = ddlAccType_Search.SelectedValue;
                    rowTemp["Acc_Man"] = vLoginID;
                    rowTemp["Rec_Date"] = vAccDateStr;
                    rowTemp["Hap_Date"] = DateTime.Today.ToShortDateString();
                    rowTemp["Source"] = vAcc_Source;
                    rowTemp["SourceID"] = vAcc_SourceID;
                    rowTemp["DocCount"] = 1;
                    rowTemp["Rate"] = 1.0;
                    rowTemp["For_Debit"] = vDebit;
                    rowTemp["For_Credit"] = vCredit;
                    rowTemp["For_Type"] = "TWD";
                    rowTemp["Stk_Money"] = 0;
                    rowTemp["Print_CNT"] = 0;
                    dtTarget.Rows.Add(rowTemp);
                }
            }
            //如果暫存資料表還有資料
            if (dtTarget.Rows.Count > 0)
            {
                //先檢查傳票借貸是否平衡
                if (vTotalDebit != vTotalCredit)
                {
                    vErrorMessage += "EXCEL 傳票：[" + vNoString.Trim() + "] 明細借貸不平！" + Environment.NewLine;
                }
                //檢查現收或現支傳票是不是有至少一筆現金科目
                if ((ddlAccType_Search.SelectedValue != "1") && (vSubject1110 == 0))
                {
                    vErrorMessage += ddlAccType_Search.SelectedItem.Text.Trim() + "傳票缺少必要的現金科目！" + Environment.NewLine;
                }
                //確定這張傳票沒有錯誤訊息時才做同步的動作
                if ((vErrorMessage == null) || (vErrorMessage.Length == 0))
                {
                    //先把日期欄位的字串轉成沒有分隔符號的民國年月日字串
                    vTempDate = DateTime.Parse(vAccDateStr);
                    vAccountYM = (vTempDate.Year - 1911).ToString("D3") + vTempDate.Month.ToString("D2") + vTempDate.Day.ToString("D2") + vTypeStr;
                    //取回同一個日期字串的最後一張傳票的號碼
                    vSQLStr_Temp = "select Max([No]) MaxNo from Account where [No] like '" + vAccountYM + "%' ";
                    vNoMax = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    //計算新的傳票號碼
                    if (Int32.TryParse(vNoMax.Replace(vAccountYM, ""), out vItemMax))
                    {
                        vItemMax++;
                    }
                    else
                    {
                        vItemMax = 1;
                    }
                    vAccountNo = vAccountYM + vItemMax.ToString("D3");
                    for (int TableRowCount = 0; TableRowCount < dtTarget.Rows.Count; TableRowCount++)
                    {
                        dtTarget.Rows[TableRowCount]["No"] = vAccountNo;
                    }
                    using (SqlBulkCopy sbkTemp = new SqlBulkCopy(vConnStr, SqlBulkCopyOptions.UseInternalTransaction))
                    {
                        try
                        {
                            sbkTemp.DestinationTableName = "Account";
                            sbkTemp.WriteToServer(dtTarget);
                            dtTarget.Clear();
                        }
                        catch (Exception eMessage)
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('" + eMessage.Message + "')");
                            Response.Write("</" + "Script>");
                        }
                    }
                }
                vItems = 0;
                vTotalCredit = 0;
                vTotalDebit = 0;
                vSubject1110 = 0;
            }
        }

        private void ImportData_X()
        {
            //Excel 2010--
            string vSQLStr_Temp = "";
            DateTime vAcc_Date_Temp = DateTime.Parse(eAccDate_Search.Text.Trim());
            DateTime vAcc_Date_Source;
            DateTime vAcc_Date_Close;
            DateTime vTempDate;
            string vAccountNo = "";
            string vAccountYM = "";
            string vNoString = "";
            string vSubject = "";
            int vDebit = 0;
            int vCredit = 0;
            int vTotalDebit = 0;
            int vTotalCredit = 0;
            string vDepNo = "";
            string vDepName = "";
            string vRemark = "";
            string vMemo2 = "";
            string vStoreNo = "";
            string vAccDateStr = "";
            int vSheetCount = 0;
            string vClass1 = "";
            int vItemMax = 0;
            int vItems = 0;
            string vNoMax = "";
            string vAcc_Source = "";
            string vAcc_SourceID = "";
            string vTypeStr = "";
            string vTempStr = "";
            int vSubject1110 = 0;
            int vCellIndex = 0;

            //取回關帳日期
            vSQLStr_Temp = "select Desscription from AC_CLS where PK = '關帳' and Class = '關帳日期' ";
            vAcc_Date_Close = DateTime.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "Desscription"));

            XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
            vSheetCount = wbExcel_X.NumberOfSheets;
            dtTarget = CreateTempDataTable();
            XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
            XSSFRow vFirstRow = (XSSFRow)sheetExcel_X.GetRow(sheetExcel_X.FirstRowNum);
            vNoString = vFirstRow.GetCell(0).ToString().Trim();
            switch (ddlAccType_Search.SelectedValue)
            {
                default:
                    vTypeStr = "0";
                    break;

                case "2":
                    vTypeStr = "4";
                    break;

                case "3":
                    vTypeStr = "5";
                    break;
            }
            for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
            {
                XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                vTempStr = "";

                if (vRowTemp_X.GetCell(0).ToString().Trim() != vNoString)
                {
                    //NoString 有變動，表示換了一張傳票，所以要把前一張傳票的資料同步進資料庫
                    //先檢查傳票借貸是否平衡
                    if ((vTotalDebit != vTotalCredit) && (dtTarget.Rows.Count > 0))
                    {
                        vErrorMessage += "EXCEL 傳票：[" + vNoString.Trim() + "] 明細借貸不平！" + Environment.NewLine;
                    }
                    //檢查現收或現支傳票是不是有至少一筆現金科目
                    if (((ddlAccType_Search.SelectedValue != "1") && (vSubject1110 == 0)) && (dtTarget.Rows.Count > 0))
                    {
                        vErrorMessage += ddlAccType_Search.SelectedItem.Text.Trim() + "傳票缺少必要的現金科目！" + Environment.NewLine;
                    }
                    //確定這張傳票沒有錯誤訊息時才做同步的動作
                    if (((vErrorMessage == null) || (vErrorMessage.Length == 0)) && (dtTarget.Rows.Count > 0))
                    {
                        //先把日期欄位的字串轉成沒有分隔符號的民國年月日字串
                        vTempDate = DateTime.Parse(vAccDateStr);
                        vAccountYM = (vTempDate.Year - 1911).ToString("D3") + vTempDate.Month.ToString("D2") + vTempDate.Day.ToString("D2") + vTypeStr;
                        //取回同一個日期字串的最後一張傳票的號碼
                        vSQLStr_Temp = "select Max([No]) MaxNo from Account where [No] like '" + vAccountYM + "%' ";
                        vNoMax = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                        //計算新的傳票號碼
                        if (Int32.TryParse(vNoMax.Replace(vAccountYM, ""), out vItemMax))
                        {
                            vItemMax++;
                        }
                        else
                        {
                            vItemMax = 1;
                        }
                        vAccountNo = vAccountYM + vItemMax.ToString("D3");
                        for (int TableRowCount = 0; TableRowCount < dtTarget.Rows.Count; TableRowCount++)
                        {
                            dtTarget.Rows[TableRowCount]["No"] = vAccountNo;
                        }
                        using (SqlBulkCopy sbkTemp = new SqlBulkCopy(vConnStr, SqlBulkCopyOptions.UseInternalTransaction))
                        {
                            try
                            {
                                sbkTemp.DestinationTableName = "Account";
                                sbkTemp.WriteToServer(dtTarget);
                                dtTarget.Clear();
                            }
                            catch (Exception eMessage)
                            {
                                Response.Write("<Script language='Javascript'>");
                                Response.Write("alert('" + eMessage.Message + "')");
                                Response.Write("</" + "Script>");
                            }
                        }
                    }
                    //同步完之後更新 NoString 的值
                    vNoString = vRowTemp_X.GetCell(0).ToString().Trim();
                    vItems = 0;
                    vTotalCredit = 0;
                    vTotalDebit = 0;
                    vSubject1110 = 0;
                }
                if (vNoString != "傳票單號")
                //if (vRowTemp_X.GetCell(0).ToString().Trim() != "傳票單號")
                {
                    //2023.10.02 因應測試回饋修改
                    //vAccDateStr = (eIgoneExcelDate_Search.Checked == true) ? vAcc_Date_Temp.ToShortDateString() : vRowTemp_X.GetCell(16).ToString().Trim();
                    /*
                    vTempDate = (eIgoneExcelDate_Search.Checked == true) ? vAcc_Date_Temp : DateTime.Parse(vRowTemp_X.GetCell(16).ToString().Trim());
                    vAccDateStr = vTempDate.ToString("yyyy/MM/dd") + " 00:00:00";
                    vSubject = vRowTemp_X.GetCell(2).ToString().Trim();
                    vDebit = (int)vRowTemp_X.GetCell(4).NumericCellValue;
                    vCredit = (int)vRowTemp_X.GetCell(5).NumericCellValue;
                    vDepNo = vRowTemp_X.GetCell(6).ToString().Trim();
                    vRemark = vRowTemp_X.GetCell(8).ToString().Trim();
                    vMemo2 = vRowTemp_X.GetCell(9).ToString().Trim();
                    vStoreNo = vRowTemp_X.GetCell(10).ToString().Trim();
                    vAcc_Source = (vRowTemp_X.GetCell(22).ToString().Trim() != "") ? vRowTemp_X.GetCell(22).ToString().Trim() : String.Empty;
                    vAcc_SourceID = (vRowTemp_X.GetCell(23).ToString().Trim() != "") ? vRowTemp_X.GetCell(23).ToString().Trim() : String.Empty; //*/
                    //2023.10.19 因為讀取 EXCEL 資料的時候經常出現沒有資料的空格被跳過，所以試一下別的做法
                    vSubject = "";
                    vDebit = 0;
                    vCredit = 0;
                    vDepNo = "";
                    vRemark = "";
                    vMemo2 = "";
                    vStoreNo = "";
                    vAccDateStr = "";
                    vAcc_Source = "";
                    vAcc_SourceID = "";
                    for (int iColumnIndex = 0; iColumnIndex < vRowTemp_X.Cells.Count; iColumnIndex++)
                    {
                        vCellIndex = vRowTemp_X.Cells[iColumnIndex].ColumnIndex;
                        switch (vCellIndex)
                        {
                            case 2:
                                vSubject = vRowTemp_X.GetCell(vCellIndex).ToString().Trim();
                                break;
                            case 4:
                                vDebit = (int)vRowTemp_X.GetCell(vCellIndex).NumericCellValue;
                                break;
                            case 5:
                                vCredit = (int)vRowTemp_X.GetCell(vCellIndex).NumericCellValue;
                                break;
                            case 6:
                                vDepNo = vRowTemp_X.GetCell(vCellIndex).ToString().Trim();
                                break;
                            case 8:
                                vRemark = vRowTemp_X.GetCell(vCellIndex).ToString().Trim();
                                break;
                            case 9:
                                vMemo2 = vRowTemp_X.GetCell(vCellIndex).ToString().Trim();
                                break;
                            case 10:
                                vStoreNo = vRowTemp_X.GetCell(vCellIndex).ToString().Trim();
                                break;
                            case 16:
                                vTempDate = (eIgoneExcelDate_Search.Checked == true) ? vAcc_Date_Temp : DateTime.Parse(vRowTemp_X.GetCell(vCellIndex).ToString().Trim());
                                vAccDateStr = vTempDate.ToString("yyyy/MM/dd") + " 00:00:00";
                                break;
                            case 22:
                                vAcc_Source = (vRowTemp_X.GetCell(vCellIndex).ToString().Trim() != "") ? vRowTemp_X.GetCell(vCellIndex).ToString().Trim() : String.Empty;
                                break;
                            case 23:
                                vAcc_SourceID = (vRowTemp_X.GetCell(vCellIndex).ToString().Trim() != "") ? vRowTemp_X.GetCell(vCellIndex).ToString().Trim() : String.Empty;
                                break;
                        }
                    }

                    //檢查傳票日期
                    if (!DateTime.TryParse(vAccDateStr, out vAcc_Date_Source))
                    {
                        vErrorMessage += "EXCEL 行數：[" + i.ToString().Trim() + "] 傳票日期有誤！" + Environment.NewLine;
                    }
                    else if (vAcc_Date_Source < vAcc_Date_Close)
                    {
                        vErrorMessage += "EXCEL 行數：[" + i.ToString().Trim() + "] 指定日期已關帳！" + Environment.NewLine;
                    }
                    //取回是否是統制科目
                    vSQLStr_Temp = "select Class1 from AC_Subject where Subject = '" + vSubject + "' ";
                    vClass1 = PF.GetValue(vConnStr, vSQLStr_Temp, "Class1");
                    if (vClass1.Trim() == "")
                    {
                        vErrorMessage += "EXCEL 行數：[" + i.ToString().Trim() + "] 找不到對應的會計科目！" + Environment.NewLine;
                    }
                    else if (vClass1 == "1")
                    {
                        vErrorMessage += "EXCEL 行數：[" + i.ToString().Trim() + "] 會計科目為統制科目！";
                    }
                    //檢查站別編號
                    if (vDepNo != "")
                    {
                        vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                        vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                        if (vDepName == "")
                        {
                            vErrorMessage += "EXCEL 行數：[" + i.ToString().Trim() + "] 站別有誤！";
                        }
                    }
                    //檢查現金科目 (111001) 內容
                    if (vSubject == "111001")
                    {
                        switch (ddlAccType_Search.SelectedValue)
                        {
                            default:
                                vTempStr = "";
                                break;

                            case "2":
                                vTempStr = (vDebit < 0) ? "EXCEL 行數：[" + i.ToString().Trim() + "] 現金收入傳票中現金科目借方金額有誤！" : "";
                                vSubject1110++;
                                break;

                            case "3":
                                vTempStr = (vCredit < 0) ? "EXCEL 行數：[" + i.ToString().Trim() + "] 現金支出傳票中現金科目貸方金額有誤！" : "";
                                vSubject1110++;
                                break;
                        }
                        if (vTempStr.Trim() != "")
                        {
                            vErrorMessage += vTempStr;
                        }
                    }
                    /*
                    //檢查金額不可為負數
                    if ((vDebit < 0) || (vCredit < 0))
                    {
                        vaErrorMessage += "EXCEL 行數：[" + i.ToString().Trim() + "] 借方或貸方金額不可為負數！";
                    } //*/

                    vItems++;
                    vTotalDebit += vDebit;
                    vTotalCredit += vCredit;
                    DataRow rowTemp = dtTarget.NewRow();
                    rowTemp["Item"] = vItems.ToString("D4");
                    rowTemp["Subject"] = vSubject;
                    rowTemp["Debit"] = vDebit;
                    rowTemp["Credit"] = vCredit;
                    rowTemp["Department"] = vDepNo;
                    rowTemp["Meno"] = vRemark;
                    rowTemp["Memo2"] = vMemo2;
                    rowTemp["Target"] = vStoreNo;
                    rowTemp["Type"] = ddlAccType_Search.SelectedValue;
                    rowTemp["Acc_Man"] = vLoginID;
                    rowTemp["Rec_Date"] = vAccDateStr;
                    rowTemp["Hap_Date"] = DateTime.Today.ToShortDateString();
                    rowTemp["Source"] = vAcc_Source;
                    rowTemp["SourceID"] = vAcc_SourceID;
                    rowTemp["DocCount"] = 1;
                    rowTemp["Rate"] = 1.0;
                    rowTemp["For_Debit"] = vDebit;
                    rowTemp["For_Credit"] = vCredit;
                    rowTemp["For_Type"] = "TWD";
                    rowTemp["Stk_Money"] = 0;
                    rowTemp["Print_CNT"] = 0;
                    dtTarget.Rows.Add(rowTemp);
                }
            }
            //如果暫存資料表還有資料
            if (dtTarget.Rows.Count > 0)
            {
                //先檢查傳票借貸是否平衡
                if (vTotalDebit != vTotalCredit)
                {
                    vErrorMessage += "EXCEL 傳票：[" + vNoString.Trim() + "] 明細借貸不平！" + Environment.NewLine;
                }
                //檢查現收或現支傳票是不是有至少一筆現金科目
                if ((ddlAccType_Search.SelectedValue != "1") && (vSubject1110 == 0))
                {
                    vErrorMessage += ddlAccType_Search.SelectedItem.Text.Trim() + "傳票缺少必要的現金科目！" + Environment.NewLine;
                }
                //確定這張傳票沒有錯誤訊息時才做同步的動作
                if ((vErrorMessage == null) || (vErrorMessage.Length == 0))
                {
                    //先把日期欄位的字串轉成沒有分隔符號的民國年月日字串
                    vTempDate = DateTime.Parse(vAccDateStr);
                    vAccountYM = (vTempDate.Year - 1911).ToString("D3") + vTempDate.Month.ToString("D2") + vTempDate.Day.ToString("D2") + vTypeStr;
                    //取回同一個日期字串的最後一張傳票的號碼
                    vSQLStr_Temp = "select Max([No]) MaxNo from Account where [No] like '" + vAccountYM + "%' ";
                    vNoMax = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    //計算新的傳票號碼
                    if (Int32.TryParse(vNoMax.Replace(vAccountYM, ""), out vItemMax))
                    {
                        vItemMax++;
                    }
                    else
                    {
                        vItemMax = 1;
                    }
                    vAccountNo = vAccountYM + vItemMax.ToString("D3");
                    for (int TableRowCount = 0; TableRowCount < dtTarget.Rows.Count; TableRowCount++)
                    {
                        dtTarget.Rows[TableRowCount]["No"] = vAccountNo;
                    }
                    using (SqlBulkCopy sbkTemp = new SqlBulkCopy(vConnStr, SqlBulkCopyOptions.UseInternalTransaction))
                    {
                        try
                        {
                            sbkTemp.DestinationTableName = "Account";
                            sbkTemp.WriteToServer(dtTarget);
                            dtTarget.Clear();
                        }
                        catch (Exception eMessage)
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('" + eMessage.Message + "')");
                            Response.Write("</" + "Script>");
                        }
                    }
                }
                vItems = 0;
                vTotalCredit = 0;
                vTotalDebit = 0;
                vSubject1110 = 0;
            }
            if ((vErrorMessage != null) && (vErrorMessage.Length > 0))
            {
                eErrorMessage.Text = vErrorMessage;
                plShowData.Visible = true;
            }
        }

        protected void bbImport_Click(object sender, EventArgs e)
        {
            string vExtName = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            if (fuExcel.FileName != "")
            {
                vExtName = Path.GetExtension(fuExcel.FileName);
                switch (vExtName)
                {
                    case ".xls": //Excel 97-2003
                        ImportData_H();
                        break;

                    case ".xlsx": //Excel 2010--
                        ImportData_X();
                        break;
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先指定要匯入的檔案。')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseError_Click(object sender, EventArgs e)
        {
            plShowData.Visible = false;
            eErrorMessage.Text = "";
        }
    }
}