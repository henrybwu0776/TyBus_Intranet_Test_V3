using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Web.UI;

namespace TyBus_Intranet_Test_V3
{
    public partial class EmpMeritsPrint : Page
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
        private DataTable dtMerits_Detail;

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
                        plReport.Visible = false;
                        plShow.Visible = false;
                        eListYear_Search.Text = (DateTime.Today.AddMonths(-1).Year - 1911).ToString();
                        ddlListMonth_Search.SelectedIndex = DateTime.Today.AddMonths(-1).Month - 1;
                    }
                    string vFDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eFDate_Search.ClientID;
                    string vFDateScript = "window.open('" + vFDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eFDate_Search.Attributes["onClick"] = vFDateScript;
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

        private string GetFNoStr(string fOriFNo)
        {
            string vResultStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_FNo = "select count(Distinct FNo) RCount from MeritsB where FNo = '" + fOriFNo + "' ";
            string vTempIndex = PF.GetValue(vConnStr, vSQLStr_FNo, "RCount");
            if (Int32.Parse(vTempIndex) > 0)
            {
                vResultStr = "   and FNo = '" + fOriFNo + "' " + Environment.NewLine;
            }
            else
            {
                string vOriTempStr = PF.ChangeChineseNumbers(fOriFNo);
                string vFNo_Temp = "";
                string vTempResult = "";
                vSQLStr_FNo = "select Distinct FNo from MeritsB order by FNo ";
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr_FNo, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    if (drTemp.HasRows)
                    {
                        while (drTemp.Read())
                        {
                            vFNo_Temp = PF.ChangeChineseNumbers(drTemp["FNo"].ToString());
                            if ((vFNo_Temp == vOriTempStr) && (vTempResult == ""))
                            {
                                vTempResult = "'" + drTemp["FNo"].ToString().Trim() + "' ";
                            }
                            else if ((vFNo_Temp == vOriTempStr) && (vResultStr != ""))
                            {
                                vTempResult += ", '" + drTemp["FNo"].ToString().Trim() + "' ";
                            }
                        }
                        if (vTempResult.Trim() != "")
                        {
                            vResultStr = "   and FNo in (" + vTempResult + ") " + Environment.NewLine;
                        }
                        else
                        {
                            vResultStr = "   and isnull(FNo, '') = '' " + Environment.NewLine;
                        }
                    }
                    else
                    {
                        vResultStr = "   and isnull(FNo, '') = '' " + Environment.NewLine;
                    }
                }
            }
            return vResultStr;
        }

        private string GetSelStr()
        {
            string vResultStr = "";
            string vWStr_FNo = (eFNo_Search.Text.Trim() != "") ? GetFNoStr(eFNo_Search.Text.Trim()) : "";
            string vWStr_FDate = (eFDate_Search.Text.Trim() != "") ? "   and FDate = '" + PF.GetAD(eFDate_Search.Text.Trim()) + "' " + Environment.NewLine : "";
            vResultStr = "select FNO, FDate, e.DepNo, (select [Name] from Department where DepNo = e.DepNo) DepNo_C, "+Environment.NewLine+
                         "       mb.EmpNo, (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        TITLE' and CLassNo =  e.Title) Title_C, " + Environment.NewLine +        
                         "       e.[Name], mb.BDate, mb.BFact1, mb.Merit1, mb.Merit2, mb.Merit3, mb.Demerit1, mb.Demerit2, mb.Demerit3, mb.Warning, mb.Alart, mb.WRIWarning, mb.Fcomment " + Environment.NewLine +
                         "  from MeritsB as mb left join Employee as e on e.EmpNo = mb.EmpNo " + Environment.NewLine +
                         " where isnull(FNo, '') <> '' " + Environment.NewLine +
                         vWStr_FNo +
                         vWStr_FDate +
                         " order by FNo, e.DepNo, e.Title, e.EmpNo ";
            return vResultStr;
        }

        private string GetSelStr_Main()
        {
            string vResultStr = "";
            string vWStr_FNo = (eFNo_Search.Text.Trim() != "") ? GetFNoStr(eFNo_Search.Text.Trim()) : "";
            string vWStr_FDate = (eFDate_Search.Text.Trim() != "") ? "   and FDate = '" + PF.GetAD(eFDate_Search.Text.Trim()) + "' " + Environment.NewLine : "";
            vResultStr = "select distinct FNO, e.DepNo, (select [Name] from Department where DepNo = e.DepNo) DepNo_C,mb.FDate " + Environment.NewLine +
                         "  from MeritsB as mb left join Employee as e on e.EmpNo = mb.EmpNo " + Environment.NewLine +
                         " where isnull(FNo, '') <> '' " + Environment.NewLine +
                         vWStr_FNo +
                         vWStr_FDate +
                         " order by FNo, e.DepNo ";
            return vResultStr;
        }

        private void OpenData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelStr();
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daTemp = new SqlDataAdapter(vSelectStr, vConnStr);
                connTemp.Open();
                dtMerits_Detail = new DataTable();
                daTemp.Fill(dtMerits_Detail);
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
            if (dtMerits_Detail.Rows.Count > 0)
            {
                gridMeritsList.DataSourceID = "";
                gridMeritsList.DataSource = dtMerits_Detail;
                gridMeritsList.DataBind();
                plShow.Visible = true;
            }
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            string vListYM = eListYear_Search.Text.Trim() + " 年 " + ddlListMonth_Search.SelectedValue.Trim() + " 月 ";
            OpenData();
            if (dtMerits_Detail.Rows.Count > 0)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                /*
                string vSQLStr_Main = GetSelStr_Main();
                DataTable dtMainP;
                using (SqlConnection connMainP = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daMainP = new SqlDataAdapter(vSQLStr_Main, connMainP);
                    connMainP.Open();
                    dtMainP = new DataTable();
                    daMainP.Fill(dtMainP);
                }

                ReportDataSource rdsPrint_Main = new ReportDataSource("EmpMeritsP_Main", dtMainP); //*/
                //ReportDataSource rdsPrint = new ReportDataSource("EmpMeritsP", dtMerits);
                ReportDataSource rdsPrint = new ReportDataSource("EmpMeritsP_Detail", dtMerits_Detail);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\EmpMeritsPrint.rdlc";
                //rvPrint.LocalReport.DataSources.Add(rdsPrint_Main);
                //rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司  "));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ListYM", vListYM));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", " 員工獎懲會名單"));
                rvPrint.LocalReport.Refresh();
                plReport.Visible = true;
                plShow.Visible = false;
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plShow.Visible = true;
            plReport.Visible = false;
        }
    }
}