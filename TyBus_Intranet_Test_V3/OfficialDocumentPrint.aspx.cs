using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class OfficialDocumentPrint : System.Web.UI.Page
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

                string vSQLStr = "";

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }

                    string vDocDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDocDate_Start_Search.ClientID;
                    string vDocDateScript = "window.open('" + vDocDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eDocDate_Start_Search.Attributes["onClick"] = vDocDateScript;

                    vDocDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDocDate_End_Search.ClientID;
                    vDocDateScript = "window.open('" + vDocDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eDocDate_End_Search.Attributes["onClick"] = vDocDateScript;

                    vDocDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eStoreDate_Start_Search.ClientID;
                    vDocDateScript = "window.open('" + vDocDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eStoreDate_Start_Search.Attributes["onClick"] = vDocDateScript;

                    vDocDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eStoreDate_End_Search.ClientID;
                    vDocDateScript = "window.open('" + vDocDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eStoreDate_End_Search.Attributes["onClick"] = vDocDateScript;

                    if (!IsPostBack)
                    {
                        eDocYears_Search.Text = (DateTime.Today.Year - 1911).ToString();
                        vSQLStr = "select count(Items) RCount from WebPermissionB where ControlName = 'OfficialDocumentGetIn' and EmpNo = '" + vLoginID + "' ";
                        int vHasPermission = Int32.Parse(PF.GetValue(vConnStr, vSQLStr, "RCount"));
                        if (vHasPermission != 0)
                        {
                            eDocDep_Search.Enabled = true;
                            eDocDep_Search.Text = "";
                            GetDocFirstWord("");
                        }
                        else
                        {
                            eDocDep_Search.Enabled = false;
                            eDocDep_Search.Text = vLoginDepNo;
                            eDocDepName_Search.Text = (string)Session["DepName"];
                            GetDocFirstWord(vLoginDepNo);
                        }
                        rvPrint.LocalReport.DataSources.Clear();
                        rvPrint.Visible = false;
                        plShowData.Visible = false;
                    }
                    else
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

        private void GetDocFirstWord(string fDepNo)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                ddlDocFirstWord_Search.Items.Clear();
                if (fDepNo == "")
                {
                    vSQLStr = "select cast('' as nvarchar) as FWNo, cast('' as nvarchar) as DocFirstCWord " + Environment.NewLine +
                              "union all " + Environment.NewLine +
                              "select FWNo, DocFirstCWord from DOCFirstWord where 1 = 1 order by FWNo";
                }
                else
                {
                    vSQLStr = "select cast('' as nvarchar) as FWNo, cast('' as nvarchar) as DocFirstCWord " + Environment.NewLine +
                              "union all " + Environment.NewLine +
                              "select FWNo, DocFirstCWord from DOCFirstWord where DepNo = '" + fDepNo + "' order by FWNo";
                }
                SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    while (drTemp.Read())
                    {
                        ListItem liTemp = new ListItem(drTemp["DocFirstCWord"].ToString().Trim(), drTemp["FWNo"].ToString().Trim());
                        ddlDocFirstWord_Search.Items.Add(liTemp);
                    }
                }
            }
        }

        private void DocumentDataBind()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            string vWStr_Date = ((eDocDate_Start_Search.Text.Trim() != "") && (eDocDate_End_Search.Text.Trim() != "")) ?
                                " and DocDate between '" + PF.TransDateString(eDocDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eDocDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                ((eDocDate_Start_Search.Text.Trim() != "") && (eDocDate_End_Search.Text.Trim() == "")) ?
                                " and DocDate = '" + PF.TransDateString(eDocDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                ((eDocDate_Start_Search.Text.Trim() == "") && (eDocDate_End_Search.Text.Trim() != "")) ?
                                " and DocDate = '" + PF.TransDateString(eDocDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine : "";

            string vWStr_StoreDate = ((eStoreDate_Start_Search.Text.Trim() != "") && (eStoreDate_End_Search.Text.Trim() != "")) ?
                                " and StoreDate between '" + PF.TransDateString(eStoreDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eStoreDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                ((eStoreDate_Start_Search.Text.Trim() != "") && (eStoreDate_End_Search.Text.Trim() == "")) ?
                                " and StoreDate = '" + PF.TransDateString(eStoreDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                ((eStoreDate_Start_Search.Text.Trim() == "") && (eStoreDate_End_Search.Text.Trim() != "")) ?
                                " and StoreDate = '" + PF.TransDateString(eStoreDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine : "";

            string vWStr_DocFirstWord = (eDocFirstWord_Search.Text.Trim() != "") ? " and DocFirstWord = '" + eDocFirstWord_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_DocNo = ((eDocNo_Start_Search.Text.Trim() != "") && (eDocNo_End_Search.Text.Trim() != "")) ? " and DocNo between '" + Int32.Parse(eDocNo_Start_Search.Text.Trim()).ToString("D4") + "' and '" + Int32.Parse(eDocNo_End_Search.Text.Trim()).ToString("D4") + "' " + Environment.NewLine :
                                 ((eDocNo_Start_Search.Text.Trim() != "") && (eDocNo_End_Search.Text.Trim() == "")) ? " and DocNo = '" + Int32.Parse(eDocNo_Start_Search.Text.Trim()).ToString("D4") + "' " + Environment.NewLine :
                                 ((eDocNo_Start_Search.Text.Trim() == "") && (eDocNo_End_Search.Text.Trim() != "")) ? " and DocNo = '" + Int32.Parse(eDocNo_End_Search.Text.Trim()).ToString("D4") + "' " + Environment.NewLine : "";
            string vWStr_DocYears = (eDocYears_Search.Text.Trim() != "") ? " and DocYears = '" + eDocYears_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vSelectStr = "SELECT DocIndex, DocDate, DocYears, " + Environment.NewLine +
                                "       DocDep, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = OfficialDocument.DocDep)) AS DocDep_C, " + Environment.NewLine +
                                "       DocFirstWord, (SELECT DocFirstCWord FROM DOCFirstWord WHERE (FWNo = OfficialDocument.DocFirstWord)) AS DocFirstWord_C, " + Environment.NewLine +
                                "       DocNo, DocSourceUnit, " + Environment.NewLine +
                                "       DocType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = CAST('公文收發登記' AS char(16)) + CAST('OfficialDocument' AS char(16)) + 'DocType') AND (CLASSNO = OfficialDocument.DocType)) AS DocType_C, " + Environment.NewLine +
                                "       DocTitle, " + Environment.NewLine +
                                "       Undertaker, (SELECT NAME FROM EMPLOYEE WHERE (LEAVEDAY IS NULL) AND(EMPNO = OfficialDocument.Undertaker)) AS Undertaker_C, " + Environment.NewLine +
                                "       OutsideDocFirstWord, OutsideDocNo, Attachement, Implementation, " + Environment.NewLine +
                                "       BuildMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (LEAVEDAY IS NULL) AND(EMPNO = OfficialDocument.BuildMan)) AS BuildMan_C, " + Environment.NewLine +
                                "       BuildDate, Remark, StoreDate, " + Environment.NewLine +
                                "       StoreMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (LEAVEDAY IS NULL) AND(EMPNO = OfficialDocument.StoreMan)) AS StoreMan_C, " + Environment.NewLine +
                                "       Remark_Store " + Environment.NewLine +
                                "  FROM OfficialDocument " + Environment.NewLine +
                                " where (1 = 1) " + Environment.NewLine +
                                vWStr_Date +
                                vWStr_DocFirstWord +
                                vWStr_DocNo +
                                vWStr_StoreDate +
                                vWStr_DocYears +
                                " order by DocNo";
            sdsOfficialDocumentP.SelectCommand = "";
            sdsOfficialDocumentP.SelectCommand = vSelectStr;
            gridOfficialDocumentP.DataBind();

            ReportDataSource rdsPrint = new ReportDataSource("dsOfficialDocumentP", sdsOfficialDocumentP);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\OfficialDocumentP.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.Refresh();
            rvPrint.Visible = true;
        }

        protected void eDocDep_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Search = eDocDep_Search.Text.Trim();
            string vDepName_Temp = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Search + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Search;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Search = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDocDep_Search.Text = vDepNo_Search;
            eDocDepName_Search.Text = vDepName_Temp;
            GetDocFirstWord(vDepNo_Search);
        }

        protected void ddlDocFirstWord_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDocFirstWord_Search.Text = ddlDocFirstWord_Search.SelectedValue.Trim();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            DocumentDataBind();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}