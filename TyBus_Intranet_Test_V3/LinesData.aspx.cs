using Amaterasu_Function;
using System;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class LinesData : Page
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
                DateTime vToday = DateTime.Today;
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
                    else
                    {
                        OpenData();
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
            string vWStr_LinesNo = ((eLinesNoS_Search.Text.Trim() != "") && (eLinesNoE_Search.Text.Trim() != "")) ?
                                   "   and a.LinesNo between '" + eLinesNoS_Search.Text.Trim() + "' and '" + eLinesNoE_Search.Text.Trim() + "' " + Environment.NewLine :
                                   ((eLinesNoS_Search.Text.Trim() != "") && (eLinesNoE_Search.Text.Trim() == "")) ?
                                   "   and a.LinesNo like '" + eLinesNoS_Search.Text.Trim() + "%' " + Environment.NewLine :
                                   ((eLinesNoS_Search.Text.Trim() == "") && (eLinesNoE_Search.Text.Trim() != "")) ?
                                   "   and a.LinesNo like '" + eLinesNoE_Search.Text.Trim() + "%' " + Environment.NewLine : "";

            string vWStr_TicketLinesNo = ((eTicketLinesNoS_Search.Text.Trim() != "") && (eTicketLinesNoE_Search.Text.Trim() != "")) ?
                                         "   and a.TicketLineNo between '" + eTicketLinesNoS_Search.Text.Trim() + "' and '" + eTicketLinesNoE_Search.Text.Trim() + "' " + Environment.NewLine :
                                         ((eTicketLinesNoS_Search.Text.Trim() != "") && (eTicketLinesNoE_Search.Text.Trim() == "")) ?
                                         "   and a.TicketLineNo like '" + eTicketLinesNoS_Search.Text.Trim() + "%' " + Environment.NewLine :
                                         ((eTicketLinesNoS_Search.Text.Trim() == "") && (eTicketLinesNoE_Search.Text.Trim() != "")) ?
                                         "   and a.TicketLineNo like '" + eTicketLinesNoE_Search.Text.Trim() + "%' " + Environment.NewLine : "";

            string vWStr_LinesNoM = ((eLinesNoMS_Search.Text.Trim() != "") && (eLinesNoME_Search.Text.Trim() != "")) ?
                                    "   and a.LinesNoM between '" + eLinesNoMS_Search.Text.Trim() + "' and '" + eLinesNoME_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eLinesNoMS_Search.Text.Trim() != "") && (eLinesNoME_Search.Text.Trim() == "")) ?
                                    "   and a.LinesNoM like '" + eLinesNoMS_Search.Text.Trim() + "%' " + Environment.NewLine :
                                    ((eLinesNoMS_Search.Text.Trim() == "") && (eLinesNoME_Search.Text.Trim() != "")) ?
                                    "   and a.LinesNoM like '" + eLinesNoME_Search.Text.Trim() + "%' " + Environment.NewLine : "";

            string vWStr_LCDLinesNo = ((eLCDLinesNoS_Search.Text.Trim() != "") && (eLCDLinesNoE_Search.Text.Trim() != "")) ?
                                      "   and b.LCDLinesNo between '" + eLCDLinesNoS_Search.Text.Trim() + "' and '" + eLCDLinesNoE_Search.Text.Trim() + "' " + Environment.NewLine :
                                      ((eLCDLinesNoS_Search.Text.Trim() != "") && (eLCDLinesNoE_Search.Text.Trim() == "")) ?
                                      "   and b.LCDLinesNo like '" + eLCDLinesNoS_Search.Text.Trim() + "%' " + Environment.NewLine :
                                      ((eLCDLinesNoS_Search.Text.Trim() == "") && (eLCDLinesNoE_Search.Text.Trim() != "")) ?
                                      "   and b.LCDLinesNo like '" + eLCDLinesNoE_Search.Text.Trim() + "%' " + Environment.NewLine : "";

            string vWStr_CallName = (eCallName_Search.Text.Trim() != "") ? "   and a.CallName like '%" + eCallName_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_LicenseNo = (eLicenseNo_Search.Text.Trim() != "") ? "   and a.LicenseNo like '%" + eLicenseNo_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_CarType = (eCarType_Search.Text.Trim() != "") ? "   and a.CarType = '" + eCarType_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CustNo = (eCustNo_Search.Text.Trim() != "") ? "   and a.CustNo like '%" + eCustNo_Search.Text.Trim() + "%' " + Environment.NewLine : "";

            vResultStr = "SELECT a.LinesNo, a.LineName, a.LicenseNo, a.ActualRun, a.ActualKM, a.Mins, a.FIXFILE, a.DATEE, a.IsFeed, a.callname, " + Environment.NewLine +
                         "       a.IsExtra, a.LinesNoM, a.operation, a.custno, (SELECT [Name] FROM CUSTOM WHERE(types = 'C') AND(code = a.custno)) AS CustName, " + Environment.NewLine +
                         "       a.amount, a.UnTaxAmt, a.allowanceKm, " + Environment.NewLine +
                         "       a.IsCycleLine, a.MainDep, (select [Name] from Department where DepNo = a.MainDep) MainDepName, " + Environment.NewLine +
                         "       a.cartype, (SELECT CLASSTXT FROM DBDICB WHERE(FKEY = '班次表          Runs            CARTYPE') AND(CLASSNO = a.cartype)) AS CarType_C, " + Environment.NewLine +
                         "       a.normal, a.FeedKm, a.FeedAmt1, a.FeedAmt2, a.highWay, a.TicketLineNo, a.BSDBSPKM, a.ApprovedTimes, a.Toll, " + Environment.NewLine +
                         "       a.ApprovedDepNo, a.LinesGovNo, a.LinesGOVExtNo, a.LicenseRun, a.LicenseRunSun, a.LicenseRunLV, a.LicenseKM, b.LCDLinesNo, a.FixedCarCount " + Environment.NewLine +
                         "  FROM Lines AS a left join LinesB as b on b.LinesNo = a.LinesNo " + Environment.NewLine +
                         " where isnull(a.LinesNo, '') <> '' " + Environment.NewLine +
                         vWStr_LinesNo +
                         vWStr_TicketLinesNo +
                         vWStr_LinesNoM +
                         vWStr_LCDLinesNo +
                         vWStr_CallName +
                         vWStr_LicenseNo +
                         vWStr_CarType +
                         vWStr_CustNo +
                         " order by a.LinesNo ";
            return vResultStr;
        }

        private void OpenData()
        {
            string vSelStr = GetSelectStr();
            dsLines_List.SelectCommand = vSelStr;
            gridLines_List.DataBind();
            //fvLines_Detail.DataBind();
            //fvLines_Detail.ChangeMode(FormViewMode.ReadOnly);
        }

        protected void ddlCarType_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eCarType_Search.Text = ddlCarType_Search.SelectedValue.Trim();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvLines_Detail_DataBound(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            int vIndex = 0;
            string vCaseDateURL = "";
            string vCaseDateScript = "";
            string vCustDataURL = "";
            string vCustDataScript = "";

            switch (fvLines_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Label eLinesNo_List = (Label)fvLines_Detail.FindControl("eLinesNo_List");
                    if (eLinesNo_List != null)
                    {
                        CheckBox cbHighWay_List = (CheckBox)fvLines_Detail.FindControl("cbHighWay_List");
                        Label eHighWay_List = (Label)fvLines_Detail.FindControl("eLinesNo_List");
                        cbHighWay_List.Checked = (eHighWay_List.Text.Trim().ToUpper() == "V");

                        CheckBox cbIsExtra_List = (CheckBox)fvLines_Detail.FindControl("cbIsExtra_List");
                        Label eIsExtra_List = (Label)fvLines_Detail.FindControl("eIsExtra_List");
                        cbIsExtra_List.Checked = (eIsExtra_List.Text.Trim().ToUpper() == "V");

                        CheckBox cbBSDBSPKM_List = (CheckBox)fvLines_Detail.FindControl("cbBSDBSPKM_List");
                        Label eBSDBSPKM_List = (Label)fvLines_Detail.FindControl("eBSDBSPKM_List");
                        cbBSDBSPKM_List.Checked = (eBSDBSPKM_List.Text.Trim().ToUpper() == "V");

                        CheckBox cbOperation_List = (CheckBox)fvLines_Detail.FindControl("cbOperation_List");
                        Label eOperation_List = (Label)fvLines_Detail.FindControl("eOperation_List");
                        cbOperation_List.Checked = (eOperation_List.Text.Trim().ToUpper() == "V");

                        CheckBox cbIsFeed_List = (CheckBox)fvLines_Detail.FindControl("cbIsFeed_List");
                        Label eIsFeed_List = (Label)fvLines_Detail.FindControl("eIsFeed_List");
                        cbIsFeed_List.Checked = (eIsFeed_List.Text.Trim().ToUpper() == "V");

                        CheckBox cbIsCycleLine_List = (CheckBox)fvLines_Detail.FindControl("cbIsCycleLine_List");
                        Label eIsCycleLine_List = (Label)fvLines_Detail.FindControl("eIsCycleLine_List");
                        cbIsCycleLine_List.Checked = (eIsCycleLine_List.Text.Trim().ToUpper() == "V");
                    }
                    break;
                case FormViewMode.Edit:
                    Label eLinesNo_Edit = (Label)fvLines_Detail.FindControl("eLinesNo_Edit");
                    if (eLinesNo_Edit != null)
                    {
                        CheckBox cbHighWay_Edit = (CheckBox)fvLines_Detail.FindControl("cbHighWay_Edit");
                        Label eHighWay_Edit = (Label)fvLines_Detail.FindControl("eLinesNo_Edit");
                        cbHighWay_Edit.Checked = (eHighWay_Edit.Text.Trim().ToUpper() == "V");

                        CheckBox cbIsExtra_Edit = (CheckBox)fvLines_Detail.FindControl("cbIsExtra_Edit");
                        Label eIsExtra_Edit = (Label)fvLines_Detail.FindControl("eIsExtra_Edit");
                        cbIsExtra_Edit.Checked = (eIsExtra_Edit.Text.Trim().ToUpper() == "V");

                        CheckBox cbBSDBSPKM_Edit = (CheckBox)fvLines_Detail.FindControl("cbBSDBSPKM_Edit");
                        Label eBSDBSPKM_Edit = (Label)fvLines_Detail.FindControl("eBSDBSPKM_Edit");
                        cbBSDBSPKM_Edit.Checked = (eBSDBSPKM_Edit.Text.Trim().ToUpper() == "V");

                        CheckBox cbOperation_Edit = (CheckBox)fvLines_Detail.FindControl("cbOperation_Edit");
                        Label eOperation_Edit = (Label)fvLines_Detail.FindControl("eOperation_Edit");
                        cbOperation_Edit.Checked = (eOperation_Edit.Text.Trim().ToUpper() == "V");

                        CheckBox cbIsFeed_Edit = (CheckBox)fvLines_Detail.FindControl("cbIsFeed_Edit");
                        Label eIsFeed_Edit = (Label)fvLines_Detail.FindControl("eIsFeed_Edit");
                        cbIsFeed_Edit.Checked = (eIsFeed_Edit.Text.Trim().ToUpper() == "V");

                        CheckBox cbIsCycleLine_Edit = (CheckBox)fvLines_Detail.FindControl("cbIsCycleLine_Edit");
                        Label eIsCycleLine_Edit = (Label)fvLines_Detail.FindControl("eIsCycleLine_Edit");
                        cbIsCycleLine_Edit.Checked = (eIsCycleLine_Edit.Text.Trim().ToUpper() == "V");

                        DropDownList ddlCarType_Edit = (DropDownList)fvLines_Detail.FindControl("ddlCarType_Edit");
                        Label eCarType_Edit = (Label)fvLines_Detail.FindControl("eCarType_Edit");
                        for (int i = 0; i < ddlCarType_Edit.Items.Count; i++)
                        {
                            vIndex = (eCarType_Edit.Text.Trim() == ddlCarType_Edit.Items[i].Value.Trim()) ? i : vIndex;
                        }
                        ddlCarType_Edit.SelectedIndex = vIndex;

                        TextBox eDateE_Edit = (TextBox)fvLines_Detail.FindControl("eDateE_Edit");
                        vCaseDateURL = "InputDate.aspx?TextboxID=" + eDateE_Edit.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eDateE_Edit.Attributes["onClick"] = vCaseDateScript;

                        TextBox eCustNo_Edit = (TextBox)fvLines_Detail.FindControl("eCustNo_Edit");
                        vCustDataURL = "SearchCust.aspx?TextBoxID=" + eCustNo_Edit.ClientID;
                        vCustDataScript = "window.open('" + vCustDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eCustNo_Edit.Attributes["onClick"] = vCustDataScript;
                    }
                    break;
                case FormViewMode.Insert:
                    TextBox eLinesNo_INS = (TextBox)fvLines_Detail.FindControl("eLinesNo_INS");
                    if (eLinesNo_INS != null)
                    {
                        CheckBox cbHighWay_INS = (CheckBox)fvLines_Detail.FindControl("cbHighWay_INS");
                        Label eHighWay_INS = (Label)fvLines_Detail.FindControl("eHighWay_INS");
                        cbHighWay_INS.Checked = (eHighWay_INS.Text.Trim().ToUpper() == "V");

                        CheckBox cbIsExtra_INS = (CheckBox)fvLines_Detail.FindControl("cbIsExtra_INS");
                        Label eIsExtra_INS = (Label)fvLines_Detail.FindControl("eIsExtra_INS");
                        cbIsExtra_INS.Checked = (eIsExtra_INS.Text.Trim().ToUpper() == "V");

                        CheckBox cbBSDBSPKM_INS = (CheckBox)fvLines_Detail.FindControl("cbBSDBSPKM_INS");
                        Label eBSDBSPKM_INS = (Label)fvLines_Detail.FindControl("eBSDBSPKM_INS");
                        cbBSDBSPKM_INS.Checked = (eBSDBSPKM_INS.Text.Trim().ToUpper() == "V");

                        CheckBox cbOperation_INS = (CheckBox)fvLines_Detail.FindControl("cbOperation_INS");
                        Label eOperation_INS = (Label)fvLines_Detail.FindControl("eOperation_INS");
                        cbOperation_INS.Checked = (eOperation_INS.Text.Trim().ToUpper() == "V");

                        CheckBox cbIsFeed_INS = (CheckBox)fvLines_Detail.FindControl("cbIsFeed_INS");
                        Label eIsFeed_INS = (Label)fvLines_Detail.FindControl("eIsFeed_INS");
                        cbIsFeed_INS.Checked = (eIsFeed_INS.Text.Trim().ToUpper() == "V");

                        CheckBox cbIsCycleLine_INS = (CheckBox)fvLines_Detail.FindControl("cbIsCycleLine_INS");
                        Label eIsCycleLine_INS = (Label)fvLines_Detail.FindControl("eIsCycleLine_INS");
                        cbIsCycleLine_INS.Checked = (eIsCycleLine_INS.Text.Trim().ToUpper() == "V");

                        DropDownList ddlCarType_INS = (DropDownList)fvLines_Detail.FindControl("ddlCarType_INS");
                        Label eCarType_INS = (Label)fvLines_Detail.FindControl("eCarType_INS");
                        for (int i = 0; i < ddlCarType_INS.Items.Count; i++)
                        {
                            vIndex = (eCarType_INS.Text.Trim() == ddlCarType_INS.Items[i].Value.Trim()) ? i : vIndex;
                        }
                        ddlCarType_INS.SelectedIndex = vIndex;

                        TextBox eDateE_INS = (TextBox)fvLines_Detail.FindControl("eDateE_INS");
                        vCaseDateURL = "InputDate.aspx?TextboxID=" + eDateE_INS.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eDateE_INS.Attributes["onClick"] = vCaseDateScript;

                        TextBox eCustNo_INS = (TextBox)fvLines_Detail.FindControl("eCustNo_INS");
                        vCustDataURL = "SearchCust.aspx?TextBoxID=" + eCustNo_INS.ClientID;
                        vCustDataScript = "window.open('" + vCustDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eCustNo_INS.Attributes["onClick"] = vCustDataScript;
                    }
                    break;
            }
        }

        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            Label eLinesNo_Edit = (Label)fvLines_Detail.FindControl("eLinesNo_Edit");
            if (eLinesNo_Edit != null)
            {
                string vLinesNo = eLinesNo_Edit.Text.Trim();
                TextBox eLineName_Edit = (TextBox)fvLines_Detail.FindControl("eLineName_Edit");
                string vLineName = eLineName_Edit.Text.Trim();
                TextBox eCallName_Edit = (TextBox)fvLines_Detail.FindControl("eCallName_Edit");
                string vCallName = eCallName_Edit.Text.Trim();
                TextBox eLinesNoM_Edit = (TextBox)fvLines_Detail.FindControl("eLinesNoM_Edit");
                string vLinesNoM = eLinesNoM_Edit.Text.Trim();
                TextBox eLinesGovNo_Edit = (TextBox)fvLines_Detail.FindControl("eLinesGovNo_Edit");
                string vLinesGovNo = eLinesGovNo_Edit.Text.Trim();
                TextBox eLinesGOVExtNo_Edit = (TextBox)fvLines_Detail.FindControl("eLinesGOVExtNo_Edit");
                string vLinesGOVExtNo = eLinesGOVExtNo_Edit.Text.Trim();
                Label eCarType_Edit = (Label)fvLines_Detail.FindControl("eCarType_Edit");
                string vCarType = eCarType_Edit.Text.Trim();
                TextBox eTicketLineNo_Edit = (TextBox)fvLines_Detail.FindControl("eTicketLineNo_Edit");
                string vTicketLineNo = eTicketLineNo_Edit.Text.Trim();
                TextBox eActualRun_Edit = (TextBox)fvLines_Detail.FindControl("eActualRun_Edit");
                string vActualRun = (eActualRun_Edit.Text.Trim() != "") ? eActualRun_Edit.Text.Trim() : "0";
                Label eHighWay_Edit = (Label)fvLines_Detail.FindControl("eHighWay_Edit");
                string vHighWay = eHighWay_Edit.Text.Trim();
                Label eIsExtra_Edit = (Label)fvLines_Detail.FindControl("eIsExtra_Edit");
                string vIsExtra = eIsExtra_Edit.Text.Trim();
                Label eBSDBSPKM_Edit = (Label)fvLines_Detail.FindControl("eBSDBSPKM_Edit");
                string vBSDBSPKM = eBSDBSPKM_Edit.Text.Trim();
                Label eOperation_Edit = (Label)fvLines_Detail.FindControl("eOperation_Edit");
                string vOperation = eOperation_Edit.Text.Trim();
                Label eIsFeed_Edit = (Label)fvLines_Detail.FindControl("eIsFeed_Edit");
                string vIsFeed = eIsFeed_Edit.Text.Trim();
                TextBox eActualKM_Edit = (TextBox)fvLines_Detail.FindControl("eActualKM_Edit");
                string vActualKM = (eActualKM_Edit.Text.Trim() != "") ? eActualKM_Edit.Text.Trim() : "0";
                TextBox eMins_Edit = (TextBox)fvLines_Detail.FindControl("eMins_Edit");
                string vMins = (eMins_Edit.Text.Trim() != "") ? eMins_Edit.Text.Trim() : "0";
                TextBox eAllowanceKm_Edit = (TextBox)fvLines_Detail.FindControl("eAllowanceKm_Edit");
                string vAllowanceKM = (eAllowanceKm_Edit.Text.Trim() != "") ? eAllowanceKm_Edit.Text.Trim() : "0";
                TextBox eNormal_Edit = (TextBox)fvLines_Detail.FindControl("eNormal_Edit");
                string vNormal = (eNormal_Edit.Text.Trim() != "") ? eNormal_Edit.Text.Trim() : "0";
                TextBox eFeedKm_Edit = (TextBox)fvLines_Detail.FindControl("eFeedKm_Edit");
                string vFeedKM = (eFeedKm_Edit.Text.Trim() != "") ? eFeedKm_Edit.Text.Trim() : "0";
                TextBox eFeedAMT2_Edit = (TextBox)fvLines_Detail.FindControl("eFeedAMT2_Edit");
                string vFeedAmt2 = (eFeedAMT2_Edit.Text.Trim() != "") ? eFeedAMT2_Edit.Text.Trim() : "0";
                TextBox eFeedAMT1_Edit = (TextBox)fvLines_Detail.FindControl("eFeedAMT1_Edit");
                string vFeedAmt1 = (eFeedAMT1_Edit.Text.Trim() != "") ? eFeedAMT1_Edit.Text.Trim() : "0";
                TextBox eApprovedTimes_Edit = (TextBox)fvLines_Detail.FindControl("eApprovedTimes_Edit");
                string vApprovedTimes = (eApprovedTimes_Edit.Text.Trim() != "") ? eApprovedTimes_Edit.Text.Trim() : "0";
                TextBox eToll_Edit = (TextBox)fvLines_Detail.FindControl("eToll_Edit");
                string vToll = eToll_Edit.Text.Trim();
                TextBox eApprovedDepNo_Edit = (TextBox)fvLines_Detail.FindControl("eApprovedDepNo_Edit");
                string vApprovedDepNo = eApprovedDepNo_Edit.Text.Trim();
                TextBox eCustNo_Edit = (TextBox)fvLines_Detail.FindControl("eCustNo_Edit");
                string vCustNo = eCustNo_Edit.Text.Trim();
                TextBox eAmount_Edit = (TextBox)fvLines_Detail.FindControl("eAmount_Edit");
                string vAmount = (eAmount_Edit.Text.Trim() != "") ? eAmount_Edit.Text.Trim() : "0";
                TextBox eUnTaxAmt_Edit = (TextBox)fvLines_Detail.FindControl("eUnTaxAmt_Edit");
                string vUnTaxAmt = (eUnTaxAmt_Edit.Text.Trim() != "") ? eUnTaxAmt_Edit.Text.Trim() : "0";
                TextBox eDateE_Edit = (TextBox)fvLines_Detail.FindControl("eDateE_Edit");
                string vDateE = eDateE_Edit.Text.Trim();
                TextBox eLicenseKM_Edit = (TextBox)fvLines_Detail.FindControl("eLicenseKM_Edit");
                string vLicenseKM = eLicenseKM_Edit.Text.Trim();
                TextBox eLicenseRun_Edit = (TextBox)fvLines_Detail.FindControl("eLicenseRun_Edit");
                string vLicenseRun = eLicenseRun_Edit.Text.Trim();
                TextBox eLicenseRunSun_Edit = (TextBox)fvLines_Detail.FindControl("eLicenseRunSun_Edit");
                string vLicenseRunSun = eLicenseRunSun_Edit.Text.Trim();
                TextBox eLicenseRunLV_Edit = (TextBox)fvLines_Detail.FindControl("eLicenseRunLV_Edit");
                string vLicenseRunLV = eLicenseRunLV_Edit.Text.Trim();
                Label eIsCycleLine_Edit = (Label)fvLines_Detail.FindControl("eIsCycleLine_Edit");
                string vIsCycleLine = eIsCycleLine_Edit.Text.Trim();
                TextBox eMainDep_Edit = (TextBox)fvLines_Detail.FindControl("eMainDep_Edit");
                string vMainDep = eMainDep_Edit.Text.Trim();
                TextBox eFixedCarCount_Edit = (TextBox)fvLines_Detail.FindControl("eFixedCarCount_Edit");
                string vFixedCarCount = eFixedCarCount_Edit.Text.Trim();
                TextBox eMainGovOffice_Edit = (TextBox)fvLines_Detail.FindControl("eMainGovOffice_Edit");
                string vMainGovOffice = eMainGovOffice_Edit.Text.Trim();

                string vSQLStr_Temp = "update Lines set LineName = @LineName, CallName = @CallName, LinesNoM = @LinesNoM, LinesGovNo = @LinesGovNo, LinesGOVExtNo = @LinesGOVExtNo, " + Environment.NewLine +
                                      "       CarType = @CarType, TicketLineNo = @TicketLineNo, ActualRun = @ActualRun, HighWay = @HighWay, IsEctra = @IsExtra, BSDBSPKM = @BSDBSPKM, " + Environment.NewLine +
                                      "       Operation = @Operation, IsFeed = @IsFeed, ActualKM = @ActualKM, Mins = @Mins, AllowanceKM = @AllowanceKM, Normal = @Normal, " + Environment.NewLine +
                                      "       FeedKM = @FeedKM, FeedAMT2 = @FeedAMT2, FeedAMT1 = @FeedAMT1, ApprovedTimes = @ApprovedTimes, Toll = @Toll, ApprovedDepNo = @ApprovedDepNo, " + Environment.NewLine +
                                      "       CustNo = @CustNo, Amount = @Amount, UnTaxAMT = @UnTaxAMT, DateE = @DateE, LicenseKM = @LicenseKM, LicenseRun = @LicenseRun, " + Environment.NewLine +
                                      "       LicenseRunSun = @LicenseRunSun, LicenseRunLV = @LicenseRunLV, IsCycleLine = @IsCycleLine, MainDep = @MainDep, "+Environment.NewLine+
                                      "       FixedCarCount = @FixedCarCount, MainGovOffice = @MainGovOffice " + Environment.NewLine +
                                      " where LinesNo = @LinesNo ";
                dsLines_Detail.UpdateCommand = vSQLStr_Temp;
                dsLines_Detail.UpdateParameters.Clear();
                dsLines_Detail.UpdateParameters.Add(new Parameter("LineName", DbType.String, (vLineName != "") ? vLineName : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("CallName", DbType.String, (vCallName != "") ? vCallName : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("LinesNoM", DbType.String, (vLinesNoM != "") ? vLinesNoM : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("LinesGovNo", DbType.String, (vLinesGovNo != "") ? vLinesGovNo : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("LinesGOVExtNo", DbType.String, (vLinesGOVExtNo != "") ? vLinesGOVExtNo : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("CarType", DbType.String, (vCarType != "") ? vCarType : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("TicketLineNo", DbType.String, (vTicketLineNo != "") ? vTicketLineNo : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("ActualRun", DbType.Int32, vActualRun));
                dsLines_Detail.UpdateParameters.Add(new Parameter("HighWay", DbType.String, (vHighWay != "") ? vHighWay : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("IsExtra", DbType.String, (vIsExtra != "") ? vIsExtra : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("BSDBSPKM", DbType.String, (vBSDBSPKM != "") ? vBSDBSPKM : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("Operation", DbType.String, (vOperation != "") ? vOperation : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("IsFeed", DbType.String, (vIsFeed != "") ? vIsFeed : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("ActualKM", DbType.Double, vActualKM));
                dsLines_Detail.UpdateParameters.Add(new Parameter("Mins", DbType.Double, vMins));
                dsLines_Detail.UpdateParameters.Add(new Parameter("AllowanceKM", DbType.Double, vAllowanceKM));
                dsLines_Detail.UpdateParameters.Add(new Parameter("Normal", DbType.Double, vNormal));
                dsLines_Detail.UpdateParameters.Add(new Parameter("FeedKM", DbType.Double, vFeedKM));
                dsLines_Detail.UpdateParameters.Add(new Parameter("FeedAMT2", DbType.Double, vFeedAmt2));
                dsLines_Detail.UpdateParameters.Add(new Parameter("FeedAMT1", DbType.Double, vFeedAmt1));
                dsLines_Detail.UpdateParameters.Add(new Parameter("ApprovedTime", DbType.Double, vApprovedTimes));
                dsLines_Detail.UpdateParameters.Add(new Parameter("Toll", DbType.Double, vToll));
                dsLines_Detail.UpdateParameters.Add(new Parameter("ApprovedDepNo", DbType.String, (vApprovedDepNo != "") ? vApprovedDepNo : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("CustNo", DbType.String, (vCustNo != "") ? vCustNo : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("Amount", DbType.Double, vAmount));
                dsLines_Detail.UpdateParameters.Add(new Parameter("UnTaxAMT", DbType.Double, vUnTaxAmt));
                dsLines_Detail.UpdateParameters.Add(new Parameter("DateE", DbType.Date, (vDateE != "") ? vDateE : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("LicenseKM", DbType.Double, (vLicenseKM != "") ? vLicenseKM : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("LicenseRun", DbType.Int32, (vLicenseRun != "") ? vLicenseRun : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("LicenseRunSun", DbType.Int32, (vLicenseRunSun != "") ? vLicenseRunSun : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("LicenseRunLV", DbType.Int32, (vLicenseRunLV != "") ? vLicenseRunLV : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("IsCycleLine", DbType.String, (vIsCycleLine != "") ? vIsCycleLine : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("MainDep", DbType.String, (vMainDep != "") ? vMainDep : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("FixedCarCount", DbType.String, (vFixedCarCount != "") ? vFixedCarCount : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("MainGovOffice", DbType.String, (vMainGovOffice != "") ? vMainGovOffice : String.Empty));
                dsLines_Detail.UpdateParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                dsLines_Detail.Update();
                gridLines_List.DataBind();
                fvLines_Detail.DataBind();
                fvLines_Detail.ChangeMode(FormViewMode.ReadOnly);
            }
        }

        protected void ddlCarType_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlCarType_Edit = (DropDownList)fvLines_Detail.FindControl("ddlCarType_Edit");
            if (ddlCarType_Edit != null)
            {
                Label eCarType_Edit = (Label)fvLines_Detail.FindControl("eCarType_Edit");
                eCarType_Edit.Text = ddlCarType_Edit.SelectedValue.Trim();
            }
        }

        protected void cbHighWay_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbHighWay_Edit = (CheckBox)fvLines_Detail.FindControl("cbHighWay_Edit");
            if (cbHighWay_Edit != null)
            {
                Label eHighWay_Edit = (Label)fvLines_Detail.FindControl("eHighWay_Edit");
                eHighWay_Edit.Text = (cbHighWay_Edit.Checked) ? "V" : "X";
            }
        }

        protected void eCustNo_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eCustNo_Edit = (TextBox)fvLines_Detail.FindControl("eCustNo_Edit");
            if (eCustNo_Edit != null)
            {
                string vSQLStr_Temp = "select [Name] from Custom where Code = '" + eCustNo_Edit.Text.Trim() + "' ";
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eCustName_Edit = (Label)fvLines_Detail.FindControl("eCustName_Edit");
                eCustName_Edit.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            }
        }

        protected void cbIsExtra_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsExtra_Edit = (CheckBox)fvLines_Detail.FindControl("cbIsExtra_Edit");
            if (cbIsExtra_Edit != null)
            {
                Label eIsExtra_Edit = (Label)fvLines_Detail.FindControl("eIsExtra_Edit");
                eIsExtra_Edit.Text = (cbIsExtra_Edit.Checked) ? "V" : "X";
            }
        }

        protected void cbOperation_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbOperation_Edit = (CheckBox)fvLines_Detail.FindControl("cbOperation_Edit");
            if (cbOperation_Edit != null)
            {
                Label eOperation_Edit = (Label)fvLines_Detail.FindControl("eOperation_Edit");
                eOperation_Edit.Text = (cbOperation_Edit.Checked) ? "V" : "X";
            }
        }

        protected void cbBSDBSPKM_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbBSDBSPKM_Edit = (CheckBox)fvLines_Detail.FindControl("cbBSDBSPKM_Edit");
            if (cbBSDBSPKM_Edit != null)
            {
                Label eBSDBSPKM_Edit = (Label)fvLines_Detail.FindControl("eBSDBSPKM_Edit");
                eBSDBSPKM_Edit.Text = (cbBSDBSPKM_Edit.Checked) ? "V" : "X";
            }
        }

        protected void cbIsFeed_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsFeed_Edit = (CheckBox)fvLines_Detail.FindControl("cbIsFeed_Edit");
            if (cbIsFeed_Edit != null)
            {
                Label eIsFeed_Edit = (Label)fvLines_Detail.FindControl("eIsFeed_Edit");
                eIsFeed_Edit.Text = (cbIsFeed_Edit.Checked) ? "V" : "X";
            }
        }

        protected void cbIsCycleLine_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsCycleLine_Edit = (CheckBox)fvLines_Detail.FindControl("cbIsCycleLine_Edit");
            if (cbIsCycleLine_Edit != null)
            {
                Label eIsCycleLine_Edit = (Label)fvLines_Detail.FindControl("eIsCycleLine_Edit");
                eIsCycleLine_Edit.Text = (cbIsCycleLine_Edit.Checked) ? "V" : "X";
            }
        }

        protected void eMainDep_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eMainDep_Edit = (TextBox)fvLines_Detail.FindControl("eMainDep_Edit");
            if (eMainDep_Edit != null)
            {
                Label eMainDepName_Edit = (Label)fvLines_Detail.FindControl("eMainDepName_Edit");
                string vDepNo = eMainDep_Edit.Text.Trim();
                string vTempStr = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vTempStr, "Name");
                if (vDepName == "")
                {
                    vDepName = vDepNo;
                    vTempStr = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' order by DepNo";
                    vDepNo = PF.GetValue(vConnStr, vTempStr, "DepNo");
                }
                eMainDep_Edit.Text = vDepNo.Trim();
                eMainDepName_Edit.Text = vDepName.Trim();
            }
        }

        protected void eLinesNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eLinesNo_INS = (TextBox)fvLines_Detail.FindControl("eLinesNo_INS");
            if ((eLinesNo_INS != null) && (eLinesNo_INS.Text.Trim() == ""))
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('ERP路線編號不可空白！')");
                Response.Write("</" + "Script>");
                eLinesNo_INS.Focus();
            }
        }

        protected void ddlCarType_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlCarType_INS = (DropDownList)fvLines_Detail.FindControl("ddlCarType_INS");
            if (ddlCarType_INS != null)
            {
                Label eCarType_INS = (Label)fvLines_Detail.FindControl("eCarType_INS");
                eCarType_INS.Text = ddlCarType_INS.SelectedValue.Trim();
            }
        }

        protected void eLicenseNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eLicenseNo_INS = (TextBox)fvLines_Detail.FindControl("eLicenseNo_INS");
            if (eLicenseNo_INS != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eLicenseRun_INS = (Label)fvLines_Detail.FindControl("eLicenseRun_INS");
                Label eLicenseRunSun_INS = (Label)fvLines_Detail.FindControl("eLicenseRunSun_INS");
                Label eLicenseRunLV_INS = (Label)fvLines_Detail.FindControl("eLicenseRunLV_INS");
                Label eLicenseKM_INS = (Label)fvLines_Detail.FindControl("eLicenseKM_INS");
                string vLicenseNo = eLicenseNo_INS.Text.Trim();
                string vTempStr = "";
                string[] vaTempStr;
                if (vLicenseNo != "")
                {
                    string vSQLStr_Temp = "select (cast(LicenseRun as varchar) + ',' + cast(LicenseRunSun as varchar) + ',' + cast(LicenseRunLV as varchar) + ',' + cast(LicenseKM as varchar)) as LicenseData " + Environment.NewLine +
                                          "  from LicenseA " + Environment.NewLine +
                                          " where LicenseNo = '" + vLicenseNo + "' ";
                    vTempStr = PF.GetValue(vConnStr, vSQLStr_Temp, "LicenseData");
                    if (vTempStr.Trim() != "")
                    {
                        vaTempStr = vTempStr.Split(',');
                        eLicenseRun_INS.Text = vaTempStr[0].Trim();
                        eLicenseRunSun_INS.Text = vaTempStr[1].Trim();
                        eLicenseRunLV_INS.Text = vaTempStr[2].Trim();
                        eLicenseKM_INS.Text = vaTempStr[3].Trim();
                    }
                    else
                    {
                        eLicenseRun_INS.Text = "無資料";
                        eLicenseRunSun_INS.Text = "無資料";
                        eLicenseRunLV_INS.Text = "無資料";
                        eLicenseKM_INS.Text = "無資料";
                    }
                }
            }
        }

        protected void cbHighWay_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbHighWay_INS = (CheckBox)fvLines_Detail.FindControl("cbHighWat_INS");
            if (cbHighWay_INS != null)
            {
                Label eHighWay_INS = (Label)fvLines_Detail.FindControl("eHighWay_INS");
                eHighWay_INS.Text = (cbHighWay_INS.Checked == true) ? "V" : "X";
            }
        }

        protected void cbIsExtra_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsExtra_INS = (CheckBox)fvLines_Detail.FindControl("cbIsExtra_INS");
            if (cbIsExtra_INS != null)
            {
                Label eIsExtra_INS = (Label)fvLines_Detail.FindControl("eIsExtra_INS");
                eIsExtra_INS.Text = (cbIsExtra_INS.Checked == true) ? "V" : "X";
            }
        }

        protected void cbBSDBSPKM_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbBSDBSPKM_INS = (CheckBox)fvLines_Detail.FindControl("cbBSDBSPKM_INS");
            if (cbBSDBSPKM_INS != null)
            {
                Label eBSDBSPKM_INS = (Label)fvLines_Detail.FindControl("eBSDBSPKM_INS");
                eBSDBSPKM_INS.Text = (cbBSDBSPKM_INS.Checked == true) ? "V" : "X";
            }
        }

        protected void cbOpration_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbOperation_INS = (CheckBox)fvLines_Detail.FindControl("cbOperation_INS");
            if (cbOperation_INS != null)
            {
                Label eOperation_INS = (Label)fvLines_Detail.FindControl("eOperation_INS");
                eOperation_INS.Text = (cbOperation_INS.Checked == true) ? "V" : "X";
            }
        }

        protected void cbIsFeed_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsFeed_INS = (CheckBox)fvLines_Detail.FindControl("cbIsFeed_INS");
            if (cbIsFeed_INS != null)
            {
                Label eIsFeed_INS = (Label)fvLines_Detail.FindControl("eIsFeed_INS");
                eIsFeed_INS.Text = (cbIsFeed_INS.Checked == true) ? "V" : "X";
            }
        }

        protected void eCustNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eCustNo_INS = (TextBox)fvLines_Detail.FindControl("eCustNo_INS");
            if (eCustNo_INS != null)
            {
                string vSQLStr_Temp = "select [Name] from Custom where Code = '" + eCustNo_INS.Text.Trim() + "' ";
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eCustName_INS = (Label)fvLines_Detail.FindControl("eCustName_INS");
                eCustName_INS.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            }
        }

        protected void cbIsCycleLine_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsCycleLine_INS = (CheckBox)fvLines_Detail.FindControl("cbIsCycleLine_INS");
            if (cbIsCycleLine_INS != null)
            {
                Label eIsCycleLine_INS = (Label)fvLines_Detail.FindControl("eIsCycleLine_INS");
                eIsCycleLine_INS.Text = (cbIsCycleLine_INS.Checked) ? "V" : "X";
            }
        }

        protected void eMainDep_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eMainDep_INS = (TextBox)fvLines_Detail.FindControl("eMainDep_INS");
            if (eMainDep_INS != null)
            {
                Label eMainDepName_INS = (Label)fvLines_Detail.FindControl("eMainDepName_INS");
                string vDepNo = eMainDep_INS.Text.Trim();
                string vTempStr = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vTempStr, "Name");
                if (vDepName == "")
                {
                    vDepName = vDepNo;
                    vTempStr = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' order by DepNo";
                    vDepNo = PF.GetValue(vConnStr, vTempStr, "DepNo");
                }
                eMainDep_INS.Text = vDepNo.Trim();
                eMainDepName_INS.Text = vDepName.Trim();
            }
        }

        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            TextBox eLinesNo_INS = (TextBox)fvLines_Detail.FindControl("eLinesNo_INS");
            if ((eLinesNo_INS != null) && (eLinesNo_INS.Text.Trim() != ""))
            {
                string vLinesNo = eLinesNo_INS.Text.Trim();
                TextBox eLineName_INS = (TextBox)fvLines_Detail.FindControl("eLineName_INS");
                string vLineName = eLineName_INS.Text.Trim();
                TextBox eCallName_INS = (TextBox)fvLines_Detail.FindControl("eCallName_INS");
                string vCallName = eCallName_INS.Text.Trim();
                TextBox eLinesNoM_INS = (TextBox)fvLines_Detail.FindControl("eLinesNoM_INS");
                string vLinesNoM = eLinesNoM_INS.Text.Trim();
                TextBox eLicenseNo_INS = (TextBox)fvLines_Detail.FindControl("eLicenseNo_INS");
                string vLicenseNo = eLicenseNo_INS.Text.Trim();
                TextBox eLinesGovNo_INS = (TextBox)fvLines_Detail.FindControl("eLinesGovNo_INS");
                string vLinesGovNo = eLinesGovNo_INS.Text.Trim();
                TextBox eLinesGOVExtNo_INS = (TextBox)fvLines_Detail.FindControl("eLinesGOVExtNo_INS");
                string vLinesGOVExtNo = eLinesGOVExtNo_INS.Text.Trim();
                Label eCarType_INS = (Label)fvLines_Detail.FindControl("eCarType_INS");
                string vCarType = eCarType_INS.Text.Trim();
                TextBox eTicketLineNo_INS = (TextBox)fvLines_Detail.FindControl("eTicketLineNo_INS");
                string vTicketLineNo = eTicketLineNo_INS.Text.Trim();
                TextBox eActualRun_INS = (TextBox)fvLines_Detail.FindControl("eActualRun_INS");
                string vActualRun = eActualRun_INS.Text.Trim();
                Label eHighWay_INS = (Label)fvLines_Detail.FindControl("eHighWay_INS");
                string vHighWay = eHighWay_INS.Text.Trim();
                TextBox eActualKM_INS = (TextBox)fvLines_Detail.FindControl("eActualKM_INS");
                string vActualKM = eActualKM_INS.Text.Trim();
                TextBox eMins_INS = (TextBox)fvLines_Detail.FindControl("eMins_INS");
                string vMins = eMins_INS.Text.Trim();
                TextBox eAllowanceKM_INS = (TextBox)fvLines_Detail.FindControl("eAllowanceKM_INS");
                string vAllowanceKM = eAllowanceKM_INS.Text.Trim();
                Label eIsExtra_INS = (Label)fvLines_Detail.FindControl("eIsExtra_INS");
                string vIsExtra = eIsExtra_INS.Text.Trim();
                TextBox eNormal_INS = (TextBox)fvLines_Detail.FindControl("eNormal_INS");
                string vNormal = eNormal_INS.Text.Trim();
                TextBox eFeedKM_INS = (TextBox)fvLines_Detail.FindControl("eFeedKM_INS");
                string vFeedKM = eFeedKM_INS.Text.Trim();
                TextBox eFeedAMT2_INS = (TextBox)fvLines_Detail.FindControl("eFeedAMT2_INS");
                string vFeedAMT2 = eFeedAMT2_INS.Text.Trim();
                TextBox eFeedAMT1_INS = (TextBox)fvLines_Detail.FindControl("eFeedAMT1_INS");
                string vFeedAMT1 = eFeedAMT1_INS.Text.Trim();
                Label eBSDBSPKM_INS = (Label)fvLines_Detail.FindControl("eBSDBSPKM_INS");
                string vBSDBSPKM = eBSDBSPKM_INS.Text.Trim();
                TextBox eApprovedTimes_INS = (TextBox)fvLines_Detail.FindControl("eApprovedTimes_INS");
                string vApprovedTimes = eApprovedTimes_INS.Text.Trim();
                TextBox eToll_INS = (TextBox)fvLines_Detail.FindControl("eToll_INS");
                string vToll = eToll_INS.Text.Trim();
                TextBox eApprovedDepNo_INS = (TextBox)fvLines_Detail.FindControl("eApprovedDepNo_INS");
                string vApprovedDepNo = eApprovedDepNo_INS.Text.Trim();
                Label eOperation_INS = (Label)fvLines_Detail.FindControl("eOperation_INS");
                string vOperation = eOperation_INS.Text.Trim();
                FileUpload fuFixFile_INS = (FileUpload)fvLines_Detail.FindControl("fuFixFile_INS");
                string vFixFile = fuFixFile_INS.FileName.Trim();
                if (vFixFile != "")
                {
                    string vUploadPath = @"\\172.20.3.17\各項上傳資料\LinesData\";
                    string vUploadFileName = vUploadPath + vFixFile;
                    fuFixFile_INS.SaveAs(vUploadFileName);
                }
                Label eIsFeed_INS = (Label)fvLines_Detail.FindControl("eIsFeed_INS");
                string vIsFeed = eIsFeed_INS.Text.Trim();
                TextBox eCustNo_INS = (TextBox)fvLines_Detail.FindControl("eCustNo_INS");
                string vCustNo = eCustNo_INS.Text.Trim();
                TextBox eAmount_INS = (TextBox)fvLines_Detail.FindControl("eAmount_INS");
                string vAmount = eAmount_INS.Text.Trim();
                TextBox eUnTaxAMT_INS = (TextBox)fvLines_Detail.FindControl("eUnTaxAMT_INS");
                string vUnTaxAMT = eUnTaxAMT_INS.Text.Trim();
                TextBox eDateE_INS = (TextBox)fvLines_Detail.FindControl("eDateE_INS");
                string vDateE = eDateE_INS.Text.Trim();
                TextBox eLicenseKM_INS = (TextBox)fvLines_Detail.FindControl("eLicenseKM_INS");
                string vLicenseKM = eLicenseKM_INS.Text.Trim();
                TextBox eLicenseRun_INS = (TextBox)fvLines_Detail.FindControl("eLicenseRun_INS");
                string vLicenseRun = eLicenseRun_INS.Text.Trim();
                TextBox eLicenseRunSun_INS = (TextBox)fvLines_Detail.FindControl("eLicenseRunSun_INS");
                string vLicenseRunSun = eLicenseRunSun_INS.Text.Trim();
                TextBox eLicenseRunLV_INS = (TextBox)fvLines_Detail.FindControl("eLicenseRunLV_INS");
                string vLicenseRunLV = eLicenseRunLV_INS.Text.Trim();
                TextBox eFixedCarCount_INS = (TextBox)fvLines_Detail.FindControl("eFixedCarCount_INS");
                string vFixedCarCount = eFixedCarCount_INS.Text.Trim();
                TextBox eMainGovOffice_INS = (TextBox)fvLines_Detail.FindControl("eMainGovOffice_INS");
                string vMainGovOffice = eMainGovOffice_INS.Text.Trim();

                dsLines_Detail.InsertCommand = "insert into Lines " + Environment.NewLine +
                                               "       (LinesNo, LineName, CallName, LinesNoM, LicenseNo, LinesGovNo, LinesGOVExtNo, CarType, TicketLineNo, ActuakRun, " + Environment.NewLine +
                                               "        HighWay, ActualKM, Mins, AllowanceKM, IsExtra, Normal, FeedKM, FeedAMT1, FeedAMT2, BSDBSPKM, " + Environment.NewLine +
                                               "        ApprovedTimes, Toll, ApprovedDepNo, Operation, FixFile, IsFeed, CustNo, Amount, UnTaxAMT, DAteE, " + Environment.NewLine +
                                               "        LicenseKM, LicenseRun, LicenseRunSun, LicenseRunLV, FixedCarCount, MainGovOffice) " + Environment.NewLine +
                                               "values (@LinesNo, @LineName, @CallName, LinesNoM, @LicenseNo, @LinesGovNo, @LinesGOVExtNo, @CarType, @TicketLineNo, @ActuakRun, " + Environment.NewLine +
                                               "        @HighWay, @ActualKM, @Mins, @AllowanceKM, @IsExtra, @Normal, @FeedKM, @FeedAMT1, @FeedAMT2, @BSDBSPKM, " + Environment.NewLine +
                                               "        @ApprovedTimes, @Toll, @ApprovedDepNo, @Operation, @FixFile, @IsFeed, @CustNo, @Amount, @UnTaxAMT, @DAteE, " + Environment.NewLine +
                                               "        @LicenseKM, @LicenseRun, @LicenseRunSun, @LicenseRunLV, @FixedCarCount, @MainGovOffice)";
                dsLines_Detail.InsertParameters.Clear();
                dsLines_Detail.InsertParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                dsLines_Detail.InsertParameters.Add(new Parameter("LineName", DbType.String, (vLineName != "") ? vLineName : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("CallName", DbType.String, (vCallName != "") ? vCallName : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("LinesNoM", DbType.String, (vLinesNoM != "") ? vLinesNoM : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("LicenseNo", DbType.String, (vLicenseNo != "") ? vLicenseNo : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, (vLinesGovNo != "") ? vLinesGovNo : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("LinesGOVExtNo", DbType.String, (vLinesGOVExtNo != "") ? vLinesGOVExtNo : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("CarType", DbType.String, (vCarType != "") ? vCarType : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, (vTicketLineNo != "") ? vTicketLineNo : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("ActualRun", DbType.Int32, (vActualRun != "") ? vActualRun : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("HighWay", DbType.String, (vHighWay != "") ? vHighWay : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("ActualKM", DbType.Double, (vActualKM != "") ? vActualKM : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("Mins", DbType.Double, (vMins != "") ? vMins : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("AllowanceKM", DbType.Double, (vAllowanceKM != "") ? vAllowanceKM : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("IsExtra", DbType.String, (vIsExtra != "") ? vIsExtra : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("Normal", DbType.Double, (vNormal != "") ? vNormal : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("FeedKM", DbType.Double, (vFeedKM != "") ? vFeedKM : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("FeedAMT1", DbType.Double, (vFeedAMT1 != "") ? vFeedAMT1 : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("FeedAMT2", DbType.Double, (vFeedAMT2 != "") ? vFeedAMT2 : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("BSDBSPKM", DbType.String, (vBSDBSPKM != "") ? vBSDBSPKM : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("ApprovedTimes", DbType.Double, (vApprovedTimes != "") ? vApprovedTimes : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("Toll", DbType.Double, (vToll != "") ? vToll : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("ApprovedDepNo", DbType.String, (vApprovedDepNo != "") ? vApprovedDepNo : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("Operation", DbType.String, (vOperation != "") ? vOperation : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("FixFile", DbType.String, (vFixFile != "") ? vFixFile : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("IsFeed", DbType.String, (vIsFeed != "") ? vIsFeed : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("CustNo", DbType.String, (vCustNo != "") ? vCustNo : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("Amount", DbType.Double, (vAmount != "") ? vAmount : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("UnTaxAMT", DbType.Double, (vUnTaxAMT != "") ? vUnTaxAMT : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("DateE", DbType.Date, (vDateE != "") ? vDateE : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("LicenseKM", DbType.Double, (vLicenseKM != "") ? vLicenseKM : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("LicenseRun", DbType.Int32, (vLicenseRun != "") ? vLicenseRun : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("LicenseRunSun", DbType.Int32, (vLicenseRunSun != "") ? vLicenseRunSun : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("LicenseRunLV", DbType.Int32, (vLicenseRunLV != "") ? vLicenseRunLV : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("FixedCarCount", DbType.String, (vFixedCarCount != "") ? vFixedCarCount : String.Empty));
                dsLines_Detail.InsertParameters.Add(new Parameter("MainGovOffice", DbType.String, (vMainGovOffice != "") ? vMainGovOffice : String.Empty));
                try
                {
                    dsLines_Detail.Insert();
                    gridLines_List.DataBind();
                    fvLines_Detail.DataBind();
                    fvLines_Detail.ChangeMode(FormViewMode.ReadOnly);
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
            else if ((eLinesNo_INS != null) && (eLinesNo_INS.Text.Trim() == ""))
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('ERP路線編號不可空白！')");
                Response.Write("</" + "Script>");
                eLinesNo_INS.Focus();
            }
        }

        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            Label eLinesNo_List = (Label)fvLines_Detail.FindControl("eLinesNo_List");
            if ((eLinesNo_List != null) && (eLinesNo_List.Text.Trim() != ""))
            {
                string vLinesNo = eLinesNo_List.Text.Trim();
                dsLines_Detail.DeleteCommand = "delete Lines where LinesNo = @LinesNo";
                dsLines_Detail.DeleteParameters.Clear();
                dsLines_Detail.DeleteParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                dsLines_Detail.Delete();
                gridLines_List.DataBind();
                fvLines_Detail.DataBind();
                fvLines_Detail.ChangeMode(FormViewMode.ReadOnly);
            }
        }
    }
}