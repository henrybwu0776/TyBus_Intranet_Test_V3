using Amaterasu_Function;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class EmpDutyHours : System.Web.UI.Page
    {
        PublicFunction PF = new PublicFunction(); //加入公用程式碼參考
        private string vLoginID = "";
        private string vLoginName = "";
        private string vLoginDepNo = "";
        private string vLoginDepName = "";
        private string vLoginTitle = "";
        private string vLoginTitleName = "";
        private string vLoginEmpType = "";
        private string vComputerName = "";
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
                        Session["EmpDutyIndex"] = "";
                    }
                    EmpDutyDataBind();
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

        /// <summary>
        /// 查詢員工加班補休時數資料
        /// </summary>
        private void EmpDutyDataBind()
        {
            string vCalBeginDay = DateTime.Today.AddYears(-1).Year.ToString("D4") + "/01/01";
            string vCalEndDay = DateTime.Today.AddDays(-1).Year.ToString("D4") + "/" + DateTime.Today.AddDays(-1).ToString("MM/dd");
            string vWStr_DepNo = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? " and a.DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? " and a.DepNo = '" + eDepNo_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? " and a.DepNo = '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = ((eEmpNo_Start_Search.Text.Trim() != "") && (eEmpNo_End_Search.Text.Trim() != "")) ? " and a.EmpNo between '" + eEmpNo_Start_Search.Text.Trim() + "' and '" + eEmpNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eEmpNo_Start_Search.Text.Trim() != "") && (eEmpNo_End_Search.Text.Trim() == "")) ? " and a.EmpNo = '" + eEmpNo_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eEmpNo_Start_Search.Text.Trim() == "") && (eEmpNo_End_Search.Text.Trim() != "")) ? " and a.EmpNo = '" + eEmpNo_End_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vSelectStr = "SELECT a.DepNo, (SELECT [Name] FROM Department WHERE (DepNo = a.DepNo)) AS DepName, " + Environment.NewLine +
                                "       a.Title, (select ClassTxt from DBDICB where FKey = '員工資料        EMPLOYEE        TITLE' and ClassNo = a.Title) Title_C, " + Environment.NewLine +
                                "       a.EmpNo, a.[Name], SUM(ISNULL(b.UsableHours, 0)) AS UseableHours " + Environment.NewLine +
                                "  FROM Employee AS a LEFT OUTER JOIN EmpDutyHours AS b ON b.EmpNo = a.EmpNo AND (b.IsAllowed = 1) " + Environment.NewLine +
                                " WHERE (a.LeaveDay IS NULL) AND (ISNULL(a.DepNo, '00') <> '00') and a.EmpNo <> 'supervisor' " + Environment.NewLine +
                                //" WHERE (ISNULL(a.DepNo, '00') <> '00') and a.EmpNo <> 'supervisor' " + Environment.NewLine +
                                "   AND b.DutyType = 'T01' AND b.DutyDateStart between '" + vCalBeginDay + "' and '" + vCalEndDay + "' " + Environment.NewLine +
                                vWStr_DepNo + vWStr_EmpNo +
                                " GROUP BY a.DepNo, a.Title, a.EmpNo, a.[Name]" + Environment.NewLine +
                                " ORDER BY a.DepNo, a.Title, a.EmpNo ";
            sdsEmpDataList.SelectCommand = "";
            sdsEmpDataList.SelectCommand = vSelectStr;
            gridEmpDataList.DataBind();
            plEmpDataList.Visible = (gridEmpDataList.Rows.Count > 0);
        }

        /// <summary>
        /// 計算可補休時數
        /// </summary>
        /// <param name="fSourceNo"></param>
        /// <returns></returns>
        private void CalUsedDuty(string fSourceNo, DateTime fRealDate, bool fIsAllowed)
        {
            string vSQLStr = "";
            string vDelStr = "";
            /*
            string vCalBeginDay = DateTime.Today.AddYears(-1).Year.ToString("D4") + "/01/01";
            string vCalEndDay = DateTime.Today.Year.ToString() + "/" + DateTime.Today.ToString("MM/dd"); //*/
            string vCalBeginDay = fRealDate.AddYears(-1).Year.ToString("D4") + "/01/01";
            string vCalEndDay = fRealDate.Year.ToString() + "/" + fRealDate.ToString("MM/dd");
            double vESCHours = 0.0;
            string vDutyIndex = "";
            string vDutyIndexItems = "";
            double vUsableHours = 0.0;
            int vItems = 0;
            string vItemsStr = "";
            string vCalEmpNo = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            //取回加班/補休員工編號
            vSQLStr = "select EmpNo from EmpDutyHours where DutyIndex = '" + fSourceNo + "' ";
            vCalEmpNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");

            //先把資料庫裡原本存的換休單B檔刪掉
            using (SqlConnection connDelSource = new SqlConnection(vConnStr))
            {
                vSQLStr = "select DutyIndexItems, DutyIndex from EmpDutyHoursB where SourceIndex = '" + fSourceNo + "' order by DutyIndexItems";
                SqlCommand cmdDelSource = new SqlCommand(vSQLStr, connDelSource);
                connDelSource.Open();
                SqlDataReader drDelSource = cmdDelSource.ExecuteReader();
                if (drDelSource.HasRows)
                {
                    while (drDelSource.Read())
                    {
                        vDutyIndex = drDelSource["DutyIndex"].ToString().Trim();
                        vDutyIndexItems = drDelSource["DutyIndexItems"].ToString().Trim();
                        vDelStr = "delete EmpDutyHoursB where DutyIndexItems = '" + vDutyIndexItems + "' ";
                        PF.ExecSQL(vConnStr, vDelStr);
                    }
                }
                drDelSource.Close();
                cmdDelSource.Cancel();
            }

            //如果換休單有核准
            if (fIsAllowed)
            {
                vSQLStr = "select ESCHours from EmpDutyHours where DutyIndex = '" + fSourceNo + "' ";
                vESCHours = (PF.GetValue(vConnStr, vSQLStr, "ESCHours") != "") ? Double.Parse(PF.GetValue(vConnStr, vSQLStr, "ESCHours")) : 0.0;

                vSQLStr = "select DutyIndex, ESCHours, UsableHours from EmpDutyHours " + Environment.NewLine +
                          " where DutyType = 'T01' and DutyDateStart between '" + vCalBeginDay + "' and '" + vCalEndDay + "' " + Environment.NewLine +
                          "   and UsableHours > 0 " + Environment.NewLine +
                          "   and EmpNo = '" + vCalEmpNo + "' " + Environment.NewLine +
                          " order by DutyDateStart ";
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        if (vESCHours > 0.0)
                        {
                            vDutyIndex = drTemp["DutyIndex"].ToString().Trim();
                            vUsableHours = double.Parse(drTemp["UsableHours"].ToString().Trim());
                            vSQLStr = "select MAX(Items) MaxItems from EmpDutyHoursB where DutyIndex = '" + vDutyIndex + "' ";
                            vItemsStr = PF.GetValue(vConnStr, vSQLStr, "MaxItems");
                            vItems = (vItemsStr == "") ? 1 : Int32.Parse(vItemsStr) + 1;
                            vItemsStr = vItems.ToString("D4");
                            vDutyIndexItems = vDutyIndex + vItemsStr;
                            if (vUsableHours >= vESCHours)
                            {
                                vSQLStr = "insert into EmpDutyHoursB (DutyIndex, Items, DutyIndexItems, SourceIndex, ESCHours)" + Environment.NewLine +
                                          "values('" + vDutyIndex + "', '" + vItemsStr + "', '" + vDutyIndexItems + "', '" + fSourceNo + "', " + vESCHours.ToString() + ")";
                                vESCHours = 0.0;
                            }
                            else
                            {
                                vSQLStr = "insert into EmpDutyHoursB (DutyIndex, Items, DutyIndexItems, SourceIndex, ESCHours)" + Environment.NewLine +
                                          "values('" + vDutyIndex + "', '" + vItemsStr + "', '" + vDutyIndexItems + "', '" + fSourceNo + "', " + vUsableHours.ToString() + ")";
                                vESCHours = vESCHours - vUsableHours;
                            }
                            PF.ExecSQL(vConnStr, vSQLStr);
                        }
                    }
                }
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            EmpDutyDataBind();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvEmpDutyDataDetail_DataBound(object sender, EventArgs e)
        {
            int vRCount = 0;
            string vSQLStr = "";
            string vDutyIndex = "";

            switch (fvEmpDutyDataDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    vSQLStr = "select count(Items) RCount from EmpDutyHoursB where DutyIndex = '" + vDutyIndex + "' ";
                    vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr, "RCount"));
                    Button bbDelete = (Button)fvEmpDutyDataDetail.FindControl("bbDelete_List");
                    Button bbEdit = (Button)fvEmpDutyDataDetail.FindControl("bbEdit_List");
                    Label eDutyDateStart = (Label)fvEmpDutyDataDetail.FindControl("eDutyDateStart_List");
                    if (eDutyDateStart != null)
                    {
                        DateTime vDutyDateStart = DateTime.Parse(eDutyDateStart.Text.Trim());
                        TimeSpan tsDayDiff = DateTime.Today - vDutyDateStart;
                        bbDelete.Enabled = ((vRCount == 0) && (tsDayDiff.Days < 365));
                        bbEdit.Enabled = ((vRCount == 0) && (tsDayDiff.Days < 365));
                        //CheckBox 的處理
                        CheckBox cbIsDisOneHour_List = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsDisOneHour_List");
                        Label eIsDisOneHour_List = (Label)fvEmpDutyDataDetail.FindControl("eIsDisOneHour_List");
                        //cbIsDisOneHour_List.Checked = (eIsDisOneHour_List.Text.Trim() == "1");
                        cbIsDisOneHour_List.Checked = (eIsDisOneHour_List.Text.Trim().ToLower() == "true");

                        CheckBox cbIsAllowed_List = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsAllowed_List");
                        Label eIsAllowed_List = (Label)fvEmpDutyDataDetail.FindControl("eIsAllowed_List");
                        //cbIsAllowed_List.Checked = (eIsAllowed_List.Text.Trim() == "1");
                        cbIsAllowed_List.Checked = (eIsAllowed_List.Text.Trim().ToLower() == "true");

                        CheckBox cbIsRejected_List = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsRejected_List");
                        Label eIsRejected_List = (Label)fvEmpDutyDataDetail.FindControl("eIsRejected_List");
                        //cbIsRejected_List.Checked = (eIsRejected_List.Text.Trim() == "1");
                        cbIsRejected_List.Checked = (eIsRejected_List.Text.Trim().ToLower() == "true");
                    }
                    break;

                case FormViewMode.Edit:
                    //建立日期輸入畫面
                    TextBox tbDutyDateStart_Edit = (TextBox)fvEmpDutyDataDetail.FindControl("eDutyDateStart_Edit");
                    string vDateTempURL_Edit = "InputDate_ChineseYears.aspx?TextboxID=" + tbDutyDateStart_Edit.ClientID;
                    string vDateTempScript_Edit = "window.open('" + vDateTempURL_Edit + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    tbDutyDateStart_Edit.Attributes["OnClick"] = vDateTempScript_Edit;
                    TextBox tbDutyDateEnd_Edit = (TextBox)fvEmpDutyDataDetail.FindControl("eDutyDateEnd_Edit");
                    vDateTempURL_Edit = "InputDate_ChineseYears.aspx?TextboxID=" + tbDutyDateEnd_Edit.ClientID;
                    vDateTempScript_Edit = "window.open('" + vDateTempURL_Edit + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    tbDutyDateEnd_Edit.Attributes["OnClick"] = vDateTempScript_Edit;
                    //CheckBox 的處理
                    CheckBox cbIsDisOneHour_Edit = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsDisOneHour_Edit");
                    Label eIsDisOneHour_Edit = (Label)fvEmpDutyDataDetail.FindControl("eIsDisOneHour_Edit");
                    //cbIsDisOneHour_Edit.Checked = (eIsDisOneHour_Edit.Text.Trim() == "1");
                    cbIsDisOneHour_Edit.Checked = (eIsDisOneHour_Edit.Text.Trim().ToLower() == "true");

                    CheckBox cbIsAllowed_Edit = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsAllowed_Edit");
                    Label eIsAllowed_Edit = (Label)fvEmpDutyDataDetail.FindControl("eIsAllowed_Edit");
                    //cbIsAllowed_Edit.Checked = (eIsAllowed_Edit.Text.Trim() == "1");
                    cbIsAllowed_Edit.Checked = (eIsAllowed_Edit.Text.Trim().ToLower() == "true");

                    CheckBox cbIsRejected_Edit = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsRejected_Edit");
                    Label eIsRejected_Edit = (Label)fvEmpDutyDataDetail.FindControl("eIsRejected_Edit");
                    //cbIsRejected_Edit.Checked = (eIsRejected_Edit.Text.Trim() == "1");
                    cbIsRejected_Edit.Checked = (eIsRejected_Edit.Text.Trim().ToLower() == "true");
                    break;

                case FormViewMode.Insert:
                    //寫入建檔人員日期
                    Label lbBuMan_INS = (Label)fvEmpDutyDataDetail.FindControl("eBuMan_Insert");
                    Label lbBuMan_C_INS = (Label)fvEmpDutyDataDetail.FindControl("eBuMan_C_Insert");
                    Label lbBuDate_INS = (Label)fvEmpDutyDataDetail.FindControl("eBuDate_Insert");
                    lbBuMan_INS.Text = vLoginID;
                    lbBuMan_C_INS.Text = (string)Session["LoginName"];
                    lbBuDate_INS.Text = PF.TransDateString(DateTime.Today, "C");
                    //建立日期輸入畫面
                    TextBox tbDutyDateStart_INS = (TextBox)fvEmpDutyDataDetail.FindControl("eDutyDateStart_Insert");
                    string vDateTempURL_INS = "InputDate_ChineseYears.aspx?TextboxID=" + tbDutyDateStart_INS.ClientID;
                    string vDateTempScript_INS = "window.open('" + vDateTempURL_INS + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    tbDutyDateStart_INS.Attributes["OnClick"] = vDateTempScript_INS;
                    TextBox tbDutyDateEnd_Insert = (TextBox)fvEmpDutyDataDetail.FindControl("eDutyDateEnd_Insert");
                    vDateTempURL_INS = "InputDate_ChineseYears.aspx?TextboxID=" + tbDutyDateEnd_Insert.ClientID;
                    vDateTempScript_INS = "window.open('" + vDateTempURL_INS + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    tbDutyDateEnd_Insert.Attributes["OnClick"] = vDateTempScript_INS;
                    //CheckBox 的處理
                    CheckBox cbIsDisOneHour_INS = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsDisOneHour_Insert");
                    Label eIsDisOneHour_INS = (Label)fvEmpDutyDataDetail.FindControl("eIsDisOneHour_Insert");
                    //cbIsDisOneHour_INS.Checked = (eIsDisOneHour_INS.Text.Trim() == "1");
                    cbIsDisOneHour_INS.Checked = (eIsDisOneHour_INS.Text.Trim().ToLower() == "true");

                    CheckBox cbIsAllowed_INS = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsAllowed_Insert");
                    Label eIsAllowed_INS = (Label)fvEmpDutyDataDetail.FindControl("eIsAllowed_Insert");
                    //cbIsAllowed_INS.Checked = (eIsAllowed_INS.Text.Trim() == "1");
                    cbIsAllowed_INS.Checked = (eIsAllowed_INS.Text.Trim().ToLower() == "true");

                    CheckBox cbIsRejected_INS = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsRejected_Insert");
                    Label eIsRejected_INS = (Label)fvEmpDutyDataDetail.FindControl("eIsRejected_Insert");
                    //cbIsRejected_INS.Checked = (eIsRejected_INS.Text.Trim() == "1");
                    cbIsRejected_INS.Checked = (eIsRejected_INS.Text.Trim().ToLower() == "true");
                    break;
            }
        }

        /// <summary>
        /// 編輯狀態下變更加班/補休時間
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eStartTime_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDutyDate_Start = (TextBox)fvEmpDutyDataDetail.FindControl("eDutyDateStart_Edit");
            if (eDutyDate_Start != null)
            {
                TextBox eDutyDate_End = (TextBox)fvEmpDutyDataDetail.FindControl("eDutyDateEnd_Edit");
                TextBox eStartTime = (TextBox)fvEmpDutyDataDetail.FindControl("eStartTime_Edit");
                TextBox eEndTime = (TextBox)fvEmpDutyDataDetail.FindControl("eEndTime_Edit");
                Label eDutyHours = (Label)fvEmpDutyDataDetail.FindControl("eDutyHours_Edit");
                Label eDutyType = (Label)fvEmpDutyDataDetail.FindControl("eDutyType_Edit");
                Label eESCHours = (Label)fvEmpDutyDataDetail.FindControl("eESCHours_Edit");
                Label eUsableHours = (Label)fvEmpDutyDataDetail.FindControl("eUsableHours_Edit");
                Label eEmpNo = (Label)fvEmpDutyDataDetail.FindControl("eEmpNo_Edit");
                TextBox eRemark = (TextBox)fvEmpDutyDataDetail.FindControl("eRemark_Edit");
                CheckBox cbIsDisHours = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsDisOneHour_Edit");
                bool vIsDisOneHour = cbIsDisHours.Checked;
                //string vCalStartDate = DateTime.Today.AddYears(-1).Year.ToString() + "/01/01";
                string vCalStartDate = (DateTime.Parse(eDutyDate_Start.Text.Trim()).Year - 1).ToString() + "/01/01";
                string vCalEndDate = DateTime.Parse(eDutyDate_Start.Text.Trim()).Year.ToString() + "/12/31";

                if ((eStartTime.Text.Trim() != "") && (eEndTime.Text.Trim() != "") && (eDutyDate_Start.Text.Trim() != "") && (eDutyDate_End.Text.Trim() == ""))
                {
                    eDutyDate_End.Text = eDutyDate_Start.Text.Trim();
                }

                if ((eStartTime.Text.Trim() != "") && (eEndTime.Text.Trim() != ""))
                {
                    string vStartTime = eStartTime.Text.Trim();
                    string vEndTime = eEndTime.Text.Trim();
                    if ((eDutyType.Text.Trim() == "T02") &&
                        ((PF.TimeStrToMins(vStartTime) > PF.TimeStrToMins("17:00")) ||
                         (PF.TimeStrToMins(vEndTime) > PF.TimeStrToMins("17:00"))))
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('您輸入的換休起迄時間超過 17:00 ！請再次確認！')");
                        Response.Write("</" + "Script>");
                    }
                    int vStartMins = PF.TimeStrToMins(vStartTime);
                    //修改計算公式，不擋換休結束時間是 17:00 的情況
                    //int vEndMins = ((eDutyType.Text.Trim() == "T02") && (PF.TimeStrToMins(vEndTime) > PF.TimeStrToMins("17:00"))) ? PF.TimeStrToMins("17:00") : PF.TimeStrToMins(vEndTime);
                    int vEndMins = PF.TimeStrToMins(vEndTime);
                    int vRestMins_S = PF.TimeStrToMins("12:00");
                    int vRestMins_E = PF.TimeStrToMins("13:00");
                    double vSpanTime = 0.0;
                    double vDutyHours = 0.0;
                    double vERPHours = 0.0;
                    string vGetERPHourStr = "";
                    string vEmpNo_Edit = eEmpNo.Text.Trim();
                    DateTime vStartDate_Edit = DateTime.Parse(eDutyDate_Start.Text.Trim());
                    string vStartDate_EditStr = vStartDate_Edit.Year.ToString("D4") + "/" + vStartDate_Edit.ToString("MM/dd");

                    if ((eDutyDate_Start.Text.Trim() == eDutyDate_End.Text.Trim()) && ((vEndMins - vStartMins) < 0))
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('時間輸入錯誤！請重新確認後輸入')");
                        Response.Write("</" + "Script>");
                        eEndTime.Text = "";
                        eEndTime.Focus();
                    }
                    else
                    {
                        if (eDutyType.Text.Trim() == "T01") //加班
                        {
                            vGetERPHourStr = "select isnull(sum(Hours), 0) TotalHours from OverDuty where ApplyMan = '" + vEmpNo_Edit + "' and RealDay = '" + vStartDate_EditStr + "' and ApplyType <> '03' ";
                            vERPHours = double.Parse(PF.GetValue(vConnStr, vGetERPHourStr, "TotalHours"));
                            if (eDutyDate_Start.Text.Trim() == eDutyDate_End.Text.Trim())
                            {
                                vSpanTime = (vEndMins - vStartMins);
                                //eDutyHours.Text = Math.Round((vSpanTime / 60.0), 0, MidpointRounding.AwayFromZero).ToString();
                                //eESCHours.Text = (eESCHours.Text.Trim() == "") ? "0" : eESCHours.Text.Trim();
                            }
                            else
                            {
                                DateTime StartDT = DateTime.Parse(eDutyDate_Start.Text.Trim() + " " + eStartTime.Text.Trim());
                                DateTime EndDT = DateTime.Parse(eDutyDate_End.Text.Trim() + " " + eEndTime.Text.Trim());
                                TimeSpan SpanDT = EndDT.Subtract(StartDT);
                                vSpanTime = (SpanDT.Hours * 60.0 + SpanDT.Minutes);
                                //eDutyHours.Text = Math.Round((vSpanTime / 60.0), 0, MidpointRounding.AwayFromZero).ToString();
                                //eESCHours.Text = (eESCHours.Text.Trim() == "") ? "0" : eESCHours.Text.Trim();
                            }
                            vDutyHours = (vIsDisOneHour) ?
                                         Math.Round((vSpanTime / 60.0), 0, MidpointRounding.AwayFromZero) - vERPHours - 1 :
                                         Math.Round((vSpanTime / 60.0), 0, MidpointRounding.AwayFromZero) - vERPHours;
                            eDutyHours.Text = (vDutyHours > 0) ?
                                              vDutyHours.ToString() :
                                              "0";
                            eESCHours.Text = (eESCHours.Text.Trim() == "") ? "0" : eESCHours.Text.Trim();
                            eUsableHours.Text = Math.Round((double.Parse(eDutyHours.Text.Trim()) - double.Parse(eESCHours.Text.Trim())), 2).ToString();
                            if (double.Parse(eUsableHours.Text.Trim()) <= 0.0)
                            {
                                Response.Write("<Script language='Javascript'>");
                                Response.Write("alert('該員於 [" + vStartDate_EditStr + "] 的加班時數已經在 ERP 系統中申報加班費，不可再申請補休')");
                                Response.Write("</" + "Script>");
                                fvEmpDutyDataDetail.ChangeMode(FormViewMode.ReadOnly);
                                fvEmpDutyDataDetail.DefaultMode = FormViewMode.ReadOnly;
                                fvEmpDutyDataDetail.DataBind();
                            }
                            else
                            {
                                eRemark.Text = ((vERPHours > 0) && (vIsDisOneHour)) ? "已扣除 ERP 系統中 [" + vStartDate_EditStr + "] 加班時數 [" + vERPHours.ToString() + "] 小時及中午午休一小時" :
                                               ((vERPHours > 0) && (!vIsDisOneHour)) ? "已扣除 ERP 系統中 [" + vStartDate_EditStr + "] 加班時數 [" + vERPHours.ToString() + "] 小時" : "";
                            }
                        }
                        else if (eDutyType.Text.Trim() == "T02") //補休
                        {
                            //string vSQLStr_Temp = "select sum(UsableHours) TotalUsableHours from EmpDutyHours where EmpNo = '" + vEmpNo_Edit + "' and DutyDateStart >= '" + vCalStartDate + "' ";
                            string vSQLStr_Temp = "select sum(UsableHours) TotalUsableHours " + Environment.NewLine +
                                                  "  from EmpDutyHours " + Environment.NewLine +
                                                  " where EmpNo = '" + vEmpNo_Edit + "' " + Environment.NewLine +
                                                  "   and (DutyDateStart between '" + vCalStartDate + "' and '" + vCalEndDate + "') ";
                            int vUsableHour = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "TotalUsableHours")) * 60;
                            eDutyHours.Text = (eDutyHours.Text.Trim() == "") ? "0" : eDutyHours.Text.Trim();

                            //先考慮換休開始時間在12:00前而且結束時間在13:00後的情況
                            if ((vStartMins < vRestMins_S) && (vEndMins > vRestMins_E))
                            {
                                vSpanTime = (vEndMins - vStartMins - 60) > 480 ? 480 : (vEndMins - vStartMins - 60);
                            }
                            //換休開始時間在12:00前而且結束時間在12:00與13:00間的情況
                            else if ((vStartMins < vRestMins_S) && ((vEndMins >= vRestMins_S) && (vEndMins <= vRestMins_E)))
                            {
                                vSpanTime = vRestMins_S - vStartMins;
                            }
                            //換休開始時間跟結束時間都在12:00前
                            else if ((vStartMins < vRestMins_S) && (vEndMins <= vRestMins_S))
                            {
                                vSpanTime = vEndMins - vStartMins;
                            }
                            //換休開始時間在12:00到13:00之間而且結束時間在13:00之後
                            else if (((vStartMins >= vRestMins_S) && (vStartMins <= vRestMins_E)) && (vEndMins > vRestMins_E))
                            {
                                vSpanTime = vEndMins - vRestMins_E;
                            }
                            //換休開始時間跟結束時間都在13:00之後
                            else if ((vStartMins >= vRestMins_E) && (vEndMins > vRestMins_E))
                            {
                                vSpanTime = vEndMins - vStartMins;
                            }
                            //都不是上述情況
                            else
                            {
                                vSpanTime = 0;
                            }
                            if (vSpanTime > vUsableHour)
                            {
                                Response.Write("<Script language='Javascript'>");
                                Response.Write("alert('欲補休時數大於可休時數！')");
                                Response.Write("</" + "Script>");
                                eStartTime.Text = "";
                                eEndTime.Text = "";
                                eStartTime.Focus();
                            }
                            else
                            {
                                eESCHours.Text = Math.Round((vSpanTime / 60.0), 0, MidpointRounding.AwayFromZero).ToString();
                                eUsableHours.Text = "0";
                            }
                        }
                    }
                }
            }
        }

        protected void eInspector_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eInspector = (TextBox)fvEmpDutyDataDetail.FindControl("eInspector_Edit");
            Label eInspectorName = (Label)fvEmpDutyDataDetail.FindControl("eInspector_C_Edit");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vInspector = eInspector.Text.Trim();
            string vSQLStr = "select [Name] from Employee where LeaveDay is null and EmpNo = '" + vInspector + "' ";
            string vInspector_C = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vInspector_C == "")
            {
                vInspector_C = vInspector;
                vSQLStr = "select EmpNo from Employee where LeaveDay is null and [Name] = '" + vInspector_C + "' ";
                vInspector = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eInspector.Text = vInspector;
            eInspectorName.Text = vInspector_C;
        }

        protected void eDepNo_Insert_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo = (TextBox)fvEmpDutyDataDetail.FindControl("eDepNo_Insert");
            Label eDepName = (Label)fvEmpDutyDataDetail.FindControl("eDepName_Insert");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo.Text.Trim();
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo.Text = vDepNo_Temp;
            eDepName.Text = vDepName_Temp;
        }

        protected void eEmpNo_Insert_TextChanged(object sender, EventArgs e)
        {
            TextBox eEmpNo = (TextBox)fvEmpDutyDataDetail.FindControl("eEmpNo_Insert");
            Label eEmpName = (Label)fvEmpDutyDataDetail.FindControl("eEmpName_Insert");
            Label eEmpError = (Label)fvEmpDutyDataDetail.FindControl("EmpDataError_Insert");
            Label eDepName = (Label)fvEmpDutyDataDetail.FindControl("eDepName_Insert");
            TextBox eDepNo = (TextBox)fvEmpDutyDataDetail.FindControl("eDepNo_Insert");
            string vDepNo_Temp = eDepNo.Text.Trim();
            string vDepName_Temp = eDepName.Text.Trim();
            eEmpError.Text = "";
            eEmpError.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo = eEmpNo.Text.Trim();
            string vSQLStr = "select [Name] from Employee where LeaveDay is null and EmpNo = '" + vEmpNo + "' ";
            //string vSQLStr = "select [Name] from Employee where EmpNo = '" + vEmpNo + "' ";
            string vEmpName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSQLStr = "select EmpNo from Employee where LeaveDay is null and [Name] = '" + vEmpName + "' ";
                //vSQLStr = "select EmpNo from Employee where [Name] = '" + vEmpName + "' ";
                vEmpNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            vSQLStr = "select count(EmpNo) RCount from Employee where LeaveDay is NULL and EmpNo = '" + vEmpNo + "' and DepNo = '" + vDepNo_Temp + "' ";
            //vSQLStr = "select count(EmpNo) RCount from Employee where EmpNo = '" + vEmpNo + "' and DepNo = '" + vDepNo_Temp + "' ";
            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr, "RCount"));
            if (vRCount > 0)
            {
                eEmpNo.Text = vEmpNo;
                eEmpName.Text = vEmpName;
            }
            else
            {
                eEmpError.Text = "員工 [" + vEmpName + "] 並非單位 [" + vDepName_Temp + "] 人員";
                eEmpError.Visible = true;
                eEmpNo.Text = "";
                eEmpName.Text = "";
                eEmpNo.Focus();
            }
        }

        protected void ddlDutyType_Insert_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTemp = (DropDownList)fvEmpDutyDataDetail.FindControl("ddlDutyType_Insert");
            Label lbTemp = (Label)fvEmpDutyDataDetail.FindControl("eDutyType_Insert");
            Label lbTemp_Ratio = (Label)fvEmpDutyDataDetail.FindControl("eHoursRatio_Insert");
            TextBox eDutyDateEnd_Insert = (TextBox)fvEmpDutyDataDetail.FindControl("eDutyDateEnd_Insert");

            lbTemp.Text = ddlTemp.SelectedValue;
            if (ddlTemp.SelectedValue.Trim() == "T02")
            {
                lbTemp_Ratio.Text = "-1";
                eDutyDateEnd_Insert.Text = "";
                eDutyDateEnd_Insert.Enabled = false;
            }
            else
            {
                lbTemp_Ratio.Text = "1";
                eDutyDateEnd_Insert.Enabled = true;
            }
        }

        /// <summary>
        /// 新增狀態下加班/補休時間變更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eStartTime_Insert_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDutyDate_Start = (TextBox)fvEmpDutyDataDetail.FindControl("eDutyDateStart_Insert");
            if (eDutyDate_Start != null)
            {
                TextBox eDutyDate_End = (TextBox)fvEmpDutyDataDetail.FindControl("eDutyDateEnd_Insert");
                TextBox eStartTime = (TextBox)fvEmpDutyDataDetail.FindControl("eStartTime_Insert");
                TextBox eEndTime = (TextBox)fvEmpDutyDataDetail.FindControl("eEndTime_Insert");
                Label eDutyHours = (Label)fvEmpDutyDataDetail.FindControl("eDutyHours_Insert");
                Label eDutyType = (Label)fvEmpDutyDataDetail.FindControl("eDutyType_Insert");
                Label eESCHours = (Label)fvEmpDutyDataDetail.FindControl("eESCHours_Insert");
                Label eUsableHours = (Label)fvEmpDutyDataDetail.FindControl("eUsableHours_Insert");
                TextBox eEmpNo = (TextBox)fvEmpDutyDataDetail.FindControl("eEmpNo_Insert");
                TextBox eRemark = (TextBox)fvEmpDutyDataDetail.FindControl("eRemark_Insert");
                CheckBox cbIsDisHours = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsDisOneHour_Insert");
                bool vIsDisOneHour = cbIsDisHours.Checked;
                //string vCalStartDate = DateTime.Today.AddYears(-1).Year.ToString() + "/01/01";
                string vCalStartDate = (DateTime.Parse(eDutyDate_Start.Text.Trim()).Year - 1).ToString() + "/01/01";
                string vCalEndDate = DateTime.Parse(eDutyDate_Start.Text.Trim()).Year.ToString() + "/12/31";

                if ((eStartTime.Text.Trim() != "") && (eEndTime.Text.Trim() != "") && (eDutyDate_Start.Text.Trim() != "") && (eDutyDate_End.Text.Trim() == ""))
                {
                    eDutyDate_End.Text = eDutyDate_Start.Text.Trim();
                }

                if ((eStartTime.Text.Trim() != "") && (eEndTime.Text.Trim() != ""))
                {
                    string vStartTime = eStartTime.Text.Trim();
                    string vEndTime = eEndTime.Text.Trim();
                    //修改計算公式，補休起迄時間超過 1700 時跳視窗提醒
                    if ((eDutyType.Text.Trim() == "T02") &&
                        ((PF.TimeStrToMins(vStartTime) > PF.TimeStrToMins("17:00")) ||
                         (PF.TimeStrToMins(vEndTime) > PF.TimeStrToMins("17:00"))))
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('您輸入的換休起迄時間超過 17:00 ！請再次確認！')");
                        Response.Write("</" + "Script>");
                    }
                    int vStartMins = PF.TimeStrToMins(vStartTime);
                    //修改計算公式，不擋換休結束時間是 17:00 的情況
                    //int vEndMins = ((eDutyType.Text.Trim() == "T02") && (PF.TimeStrToMins(vEndTime) > PF.TimeStrToMins("17:00"))) ? PF.TimeStrToMins("17:00") : PF.TimeStrToMins(vEndTime);
                    int vEndMins = PF.TimeStrToMins(vEndTime);
                    int vRestMins_S = PF.TimeStrToMins("12:00");
                    int vRestMins_E = PF.TimeStrToMins("13:00");
                    double vSpanTime = 0.0;
                    double vDutyHours = 0.0;
                    double vERPHours = 0.0;
                    string vGetERPHourStr = "";
                    string vEmpNo_INS = eEmpNo.Text.Trim();
                    DateTime vStartDate_INS = DateTime.Parse(eDutyDate_Start.Text.Trim());
                    string vStartDate_INSStr = vStartDate_INS.Year.ToString("D4") + "/" + vStartDate_INS.ToString("MM/dd");

                    if ((eDutyDate_Start.Text.Trim() == eDutyDate_End.Text.Trim()) && ((vEndMins - vStartMins) < 0))
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('時間輸入錯誤！請重新確認後輸入')");
                        Response.Write("</" + "Script>");
                        eEndTime.Text = "";
                        eEndTime.Focus();
                    }
                    else
                    {
                        if (eDutyType.Text.Trim() == "T01") //加班
                        {
                            vGetERPHourStr = "select isnull(sum(Hours), 0) TotalHours from OverDuty where ApplyMan = '" + vEmpNo_INS + "' and RealDay = '" + vStartDate_INSStr + "' and ApplyType <> '03' ";
                            vERPHours = double.Parse(PF.GetValue(vConnStr, vGetERPHourStr, "TotalHours"));
                            if (eDutyDate_Start.Text.Trim() == eDutyDate_End.Text.Trim())
                            {
                                vSpanTime = (vEndMins - vStartMins);
                                //eDutyHours.Text = Math.Round((vSpanTime / 60.0), 0, MidpointRounding.AwayFromZero).ToString();
                                //eESCHours.Text = (eESCHours.Text.Trim() == "") ? "0" : eESCHours.Text.Trim();
                                //eUsableHours.Text = Math.Round((double.Parse(eDutyHours.Text.Trim()) - double.Parse(eESCHours.Text.Trim())), 0, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                DateTime StartDT = DateTime.Parse(eDutyDate_Start.Text.Trim() + " " + eStartTime.Text.Trim());
                                DateTime EndDT = DateTime.Parse(eDutyDate_End.Text.Trim() + " " + eEndTime.Text.Trim());
                                TimeSpan SpanDT = EndDT.Subtract(StartDT);
                                vSpanTime = (SpanDT.Hours * 60.0 + SpanDT.Minutes);
                                //eDutyHours.Text = Math.Round((vSpanTime / 60.0), 0, MidpointRounding.AwayFromZero).ToString();
                                //eESCHours.Text = (eESCHours.Text.Trim() == "") ? "0" : eESCHours.Text.Trim();
                                //eUsableHours.Text = Math.Round((double.Parse(eDutyHours.Text.Trim()) - double.Parse(eESCHours.Text.Trim())), 0, MidpointRounding.AwayFromZero).ToString();
                            }
                            vDutyHours = (vIsDisOneHour) ?
                                         Math.Round((vSpanTime / 60.0), 0, MidpointRounding.AwayFromZero) - vERPHours - 1 :
                                         Math.Round((vSpanTime / 60.0), 0, MidpointRounding.AwayFromZero) - vERPHours;
                            eDutyHours.Text = (vDutyHours > 0) ?
                                              vDutyHours.ToString() :
                                              "0";
                            eESCHours.Text = (eESCHours.Text.Trim() == "") ? "0" : eESCHours.Text.Trim();
                            eUsableHours.Text = Math.Round((double.Parse(eDutyHours.Text.Trim()) - double.Parse(eESCHours.Text.Trim())), 0, MidpointRounding.AwayFromZero).ToString();
                            if (double.Parse(eDutyHours.Text.Trim()) <= 0.0)
                            {
                                Response.Write("<Script language='Javascript'>");
                                Response.Write("alert('該員於 [" + vStartDate_INSStr + "] 的加班時數已經在 ERP 系統中申報加班費，不可再申請補休')");
                                Response.Write("</" + "Script>");
                                fvEmpDutyDataDetail.ChangeMode(FormViewMode.ReadOnly);
                                fvEmpDutyDataDetail.DefaultMode = FormViewMode.ReadOnly;
                                fvEmpDutyDataDetail.DataBind();
                            }
                            else
                            {
                                eRemark.Text = ((vERPHours > 0) && (vIsDisOneHour)) ? "已扣除 ERP 系統中 [" + vStartDate_INSStr + "] 加班時數 [" + vERPHours.ToString() + "] 小時及中午午休一小時" :
                                               ((vERPHours > 0) && (!vIsDisOneHour)) ? "已扣除 ERP 系統中 [" + vStartDate_INSStr + "] 加班時數 [" + vERPHours.ToString() + "] 小時" : "";
                            }
                        }
                        else if (eDutyType.Text.Trim() == "T02") //補休
                        {
                            //string vSQLStr_Temp = "select sum(UsableHours) TotalUsableHours from EmpDutyHours where EmpNo = '" + vEmpNo_INS + "' and DutyDateStart >= '" + vCalStartDate + "' ";
                            string vSQLStr_Temp = "select sum(UsableHours) TotalUsableHours " + Environment.NewLine +
                                                  "  from EmpDutyHours " + Environment.NewLine +
                                                  " where EmpNo = '" + vEmpNo_INS + "' " + Environment.NewLine +
                                                  "   and (DutyDateStart between '" + vCalStartDate + "' and '" + vCalEndDate + "') ";
                            int vUsableHour = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "TotalUsableHours")) * 60;

                            eDutyHours.Text = (eDutyHours.Text.Trim() == "") ? "0" : eDutyHours.Text.Trim();

                            //先考慮換休開始時間在12:00前而且結束時間在13:00後的情況
                            if ((vStartMins < vRestMins_S) && (vEndMins > vRestMins_E))
                            {
                                vSpanTime = (vEndMins - vStartMins - 60) > 480 ? 480 : (vEndMins - vStartMins - 60);
                            }
                            //換休開始時間在12:00前而且結束時間在12:00與13:00間的情況
                            else if ((vStartMins < vRestMins_S) && ((vEndMins >= vRestMins_S) && (vEndMins <= vRestMins_E)))
                            {
                                vSpanTime = vRestMins_S - vStartMins;
                            }
                            //換休開始時間跟結束時間都在12:00前
                            else if ((vStartMins < vRestMins_S) && (vEndMins <= vRestMins_S))
                            {
                                vSpanTime = vEndMins - vStartMins;
                            }
                            //換休開始時間在12:00到13:00之間而且結束時間在13:00之後
                            else if (((vStartMins >= vRestMins_S) && (vStartMins <= vRestMins_E)) && (vEndMins > vRestMins_E))
                            {
                                vSpanTime = vEndMins - vRestMins_E;
                            }
                            //換休開始時間跟結束時間都在13:00之後
                            else if ((vStartMins >= vRestMins_E) && (vEndMins > vRestMins_E))
                            {
                                vSpanTime = vEndMins - vStartMins;
                            }
                            //都不是上述情況
                            else
                            {
                                vSpanTime = 0;
                            }
                            if (vSpanTime > vUsableHour)
                            {
                                Response.Write("<Script language='Javascript'>");
                                Response.Write("alert('欲補休時數大於可休時數！')");
                                Response.Write("</" + "Script>");
                                eStartTime.Text = "";
                                eEndTime.Text = "";
                                eStartTime.Focus();
                            }
                            else
                            {
                                eESCHours.Text = Math.Round((vSpanTime / 60.0), 0, MidpointRounding.AwayFromZero).ToString();
                                eUsableHours.Text = "0";
                            }
                        }
                    }
                }
            }
        }

        protected void eInspector_Insert_TextChanged(object sender, EventArgs e)
        {
            TextBox eInspector = (TextBox)fvEmpDutyDataDetail.FindControl("eInspector_Insert");
            Label eInspectorName = (Label)fvEmpDutyDataDetail.FindControl("eInspector_C_Insert");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vInspector = eInspector.Text.Trim();
            string vSQLStr = "select [Name] from Employee where LeaveDay is null and EmpNo = '" + vInspector + "' ";
            string vInspector_C = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vInspector_C == "")
            {
                vInspector_C = vInspector;
                vSQLStr = "select EmpNo from Employee where LeaveDay is null and [Name] = '" + vInspector_C + "' ";
                vInspector = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eInspector.Text = vInspector;
            eInspectorName.Text = vInspector_C;
        }

        /// <summary>
        /// 刪除加班/請休單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            Label eDutyIndex = (Label)fvEmpDutyDataDetail.FindControl("eDutyIndex_List");
            Label eDutyType = (Label)fvEmpDutyDataDetail.FindControl("eDutyType_List");
            string vDutyIndex = eDutyIndex.Text.Trim();
            string vDutyType = eDutyType.Text.Trim();
            string vSQLStr = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            switch (vDutyType)
            {
                case "T01": //加班單
                    vSQLStr = "select Count(DutyIndexItems) RCount from EmpDutyHoursB where DutyIndex = '" + vDutyIndex + "' ";
                    int ItemCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr, "RCount"));
                    if (ItemCount <= 0)
                    {
                        sdsEmpDutyDataDetail.DeleteParameters.Clear();
                        sdsEmpDutyDataDetail.DeleteParameters.Add(new Parameter("DutyIndex", DbType.String, vDutyIndex));
                        sdsEmpDutyDataDetail.Delete();
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('加班單號 [ " + vDutyIndex + " ] 已經有補休抵用記錄，不可刪除！')");
                        Response.Write("</" + "Script>");
                    }
                    break;
                case "T02": //換休單
                    vSQLStr = "delete EmpDutyHoursB where SourceIndex = '" + vDutyIndex + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr);
                    sdsEmpDutyDataDetail.DeleteParameters.Clear();
                    sdsEmpDutyDataDetail.DeleteParameters.Add(new Parameter("DutyIndex", DbType.String, vDutyIndex));
                    sdsEmpDutyDataDetail.Delete();
                    break;
            }
        }

        /// <summary>
        /// 加班/請休單修改存檔
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            Label eDutyIndex_Edit = (Label)fvEmpDutyDataDetail.FindControl("eDutyIndex_Edit");
            if (eDutyIndex_Edit != null)
            {
                Label eDutyType = (Label)fvEmpDutyDataDetail.FindControl("eDutyType_Edit");
                TextBox eDutydateS_Edit = (TextBox)fvEmpDutyDataDetail.FindControl("eDutyDateStart_Edit");
                DateTime vDutyDate_S_Edit = (eDutydateS_Edit.Text.Trim() != "") ? DateTime.Parse(eDutydateS_Edit.Text.Trim()) : DateTime.Today;
                TextBox eStartTime_Edit = (TextBox)fvEmpDutyDataDetail.FindControl("eStartTime_Edit");
                TextBox eDutyDateE_Edit = (TextBox)fvEmpDutyDataDetail.FindControl("eDutyDateEnd_Edit");
                DateTime vDutyDate_E_Edit = (eDutyDateE_Edit.Text.Trim() != "") ? DateTime.Parse(eDutyDateE_Edit.Text.Trim()) : DateTime.Today;
                TextBox eEndTime_Edit = (TextBox)fvEmpDutyDataDetail.FindControl("eEndTime_Edit");
                Label eDutyHours_Edit = (Label)fvEmpDutyDataDetail.FindControl("eDutyHours_Edit");
                TextBox eAssignReason_Edit = (TextBox)fvEmpDutyDataDetail.FindControl("eAssignReason_Edit");
                TextBox eRemark_Edit = (TextBox)fvEmpDutyDataDetail.FindControl("eRemark_Edit");
                Label eESCHours_Edit = (Label)fvEmpDutyDataDetail.FindControl("eESCHours_Edit");
                Label eUsableHours_Edit = (Label)fvEmpDutyDataDetail.FindControl("eUsableHours_Edit");
                TextBox eInspector_Edit = (TextBox)fvEmpDutyDataDetail.FindControl("eInspector_Edit");
                CheckBox cbIsAllowed_Edit = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsAllowed_Edit");
                CheckBox cbIsRejected_Edit = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsRejected_Edit");
                CheckBox cbIsDisOneHour_Edit = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsDisOneHour_Edit");

                string vDutyType = eDutyType.Text.Trim();

                string vEditStr = "UPDATE EmpDutyHours " + Environment.NewLine +
                                  "   SET DutyDateStart = @DutyDateStart, StartTime = @StartTime, DutyDateEnd = @DutyDateEnd, EndTime = @EndTime, " + Environment.NewLine +
                                  "       DutyHours = @DutyHours, AssignReason = @AssignReason, Remark = @Remark, ESCHours = @ESCHours, " + Environment.NewLine +
                                  "       UsableHours = @UsableHours, Inspector = @Inspector, IsAllowed = @IsAllowed, IsRejected = @IsRejected, " + Environment.NewLine +
                                  "       IsDisOneHour = @IsDisOneHour " + Environment.NewLine +
                                  " WHERE (DutyIndex = @DutyIndex)";

                try
                {
                    if (eDutydateS_Edit.Text.Trim() != "")
                    {
                        sdsEmpDutyDataDetail.UpdateCommand = vEditStr;
                        sdsEmpDutyDataDetail.UpdateParameters.Clear();
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("DutyDateStart", DbType.Date, (eDutydateS_Edit.Text.Trim() != "") ? vDutyDate_S_Edit.ToShortDateString() : String.Empty));
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("StartTime", DbType.String, (eStartTime_Edit.Text.Trim() != "") ? eStartTime_Edit.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("DutyDateEnd", DbType.Date, (eDutyDateE_Edit.Text.Trim() != "") ? vDutyDate_E_Edit.ToShortDateString() : String.Empty));
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("EndTime", DbType.String, (eEndTime_Edit.Text.Trim() != "") ? eEndTime_Edit.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("DutyHours", DbType.Double, (eDutyHours_Edit.Text.Trim() != "") ? eDutyHours_Edit.Text.Trim() : "0"));
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("AssignReason", DbType.String, (eAssignReason_Edit.Text.Trim() != "") ? eAssignReason_Edit.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("Remark", DbType.String, (eRemark_Edit.Text.Trim() != "") ? eRemark_Edit.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("ESCHours", DbType.Double, (eESCHours_Edit.Text.Trim() != "") ? eESCHours_Edit.Text.Trim() : "0"));
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("UsableHours", DbType.Double, (eUsableHours_Edit.Text.Trim() != "") ? eUsableHours_Edit.Text.Trim() : "0"));
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("Inspector", DbType.String, (eInspector_Edit.Text.Trim() != "") ? eInspector_Edit.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("IsAllowed", DbType.Boolean, (cbIsAllowed_Edit.Checked) ? "true" : "false"));
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("IsRejected", DbType.Boolean, (cbIsRejected_Edit.Checked) ? "true" : "false"));
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("IsDisOneHour", DbType.Boolean, (cbIsDisOneHour_Edit.Checked) ? "true" : "false"));
                        sdsEmpDutyDataDetail.UpdateParameters.Add(new Parameter("DutyIndex", DbType.String, eDutyIndex_Edit.Text.Trim()));

                        sdsEmpDutyDataDetail.Update();

                        if (vDutyType == "T02")
                        {
                            //Session["EmpDutyIndex"] = eDutyIndex_Edit.Text.Trim();
                            CalUsedDuty(eDutyIndex_Edit.Text.Trim(), vDutyDate_S_Edit, cbIsAllowed_Edit.Checked);
                        }
                        else
                        {
                            //Session["EmpDutyIndex"] = "";
                        }
                        fvEmpDutyDataDetail.ChangeMode(FormViewMode.ReadOnly);
                        this.gridEmpDataList.DataBind();
                        this.gridEmpDutyDataList.DataBind();
                    }
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
        }

        /// <summary>
        /// 加班/請休單新增存檔
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Insert_Click(object sender, EventArgs e)
        {
            string vDutyIndex = "";
            TextBox eDepNo_INS = (TextBox)fvEmpDutyDataDetail.FindControl("eDepNo_Insert");
            if (eDepNo_INS != null)
            {
                Label eDepName_INS = (Label)fvEmpDutyDataDetail.FindControl("eDepName_Insert");
                TextBox eEmpNo_INS = (TextBox)fvEmpDutyDataDetail.FindControl("eEmpNo_Insert");
                Label eEmpName_INS = (Label)fvEmpDutyDataDetail.FindControl("eEmpName_Insert");
                Label eDutyType_INS = (Label)fvEmpDutyDataDetail.FindControl("eDutyType_Insert");
                TextBox eDutyDateS_INS = (TextBox)fvEmpDutyDataDetail.FindControl("eDutyDateStart_Insert");
                DateTime vDutyDate_S = (eDutyDateS_INS.Text.Trim() != "") ? DateTime.Parse(eDutyDateS_INS.Text.Trim()) : DateTime.Today;
                TextBox eStartTime_INS = (TextBox)fvEmpDutyDataDetail.FindControl("eStartTime_Insert");
                TextBox eDutyDateE_INS = (TextBox)fvEmpDutyDataDetail.FindControl("eDutyDateEnd_Insert");
                DateTime vDutyDate_E = (eDutyDateE_INS.Text.Trim() != "") ? DateTime.Parse(eDutyDateE_INS.Text.Trim()) : DateTime.Today;
                TextBox eEndTime_INS = (TextBox)fvEmpDutyDataDetail.FindControl("eEndTime_Insert");
                Label eDutyHours_INS = (Label)fvEmpDutyDataDetail.FindControl("eDutyHours_Insert");
                Label eBuMan_INS = (Label)fvEmpDutyDataDetail.FindControl("eBuMan_Insert");
                Label eBuDate_INS = (Label)fvEmpDutyDataDetail.FindControl("eBuDate_Insert");
                DateTime vBuDate = (eBuDate_INS.Text.Trim() != "") ? DateTime.Parse(eBuDate_INS.Text.Trim()) : DateTime.Today;
                TextBox eAssignReason_INS = (TextBox)fvEmpDutyDataDetail.FindControl("eAssignReason_Insert");
                TextBox eRemark_INS = (TextBox)fvEmpDutyDataDetail.FindControl("eRemark_Insert");
                TextBox eInspector_INS = (TextBox)fvEmpDutyDataDetail.FindControl("eInspector_Insert");
                CheckBox cbIsAllowed_INS = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsAllowed_Insert");
                CheckBox cbIsRejected_INS = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsRejected_Insert");
                TextBox eRejectReason_INS = (TextBox)fvEmpDutyDataDetail.FindControl("eRejectReason_Insert");
                Label eHoursRatio_INS = (Label)fvEmpDutyDataDetail.FindControl("eHoursRatio_Insert");
                Label eESCHours_INS = (Label)fvEmpDutyDataDetail.FindControl("eESCHours_Insert");
                Label eUsableHours_INS = (Label)fvEmpDutyDataDetail.FindControl("eUsableHours_Insert");
                CheckBox cbIsDisOneHour_INS = (CheckBox)fvEmpDutyDataDetail.FindControl("cbIsDisOneHour_Insert");

                string vBuildMonthStr = (vDutyDate_S.Year - 1911).ToString("D3") + vDutyDate_S.ToString("MM");
                vDutyIndex = vBuildMonthStr + "T" + eEmpNo_INS.Text.Trim();
                string vSQLStr_Temp = "select Max(DutyIndex) MaxIndex from EmpDutyHours where DutyIndex like '" + vDutyIndex + "%' ";
                string vIndexStr = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                int vIndex = (vIndexStr != "") ? Int32.Parse(vIndexStr.Replace(vDutyIndex, "")) + 1 : 1;
                vDutyIndex = vDutyIndex + vIndex.ToString("D3");
                string vDutyType = eDutyType_INS.Text.Trim();

                string vInsertStr = "INSERT INTO EmpDutyHours " + Environment.NewLine +
                                    "       (DutyIndex, DepNo, DepName, EmpNo, EmpName, DutyType, DutyDateStart, StartTime, DutyDateEnd, EndTime, " + Environment.NewLine +
                                    "        DutyHours, BuMan, BuDate, AssignReason, Remark, Inspector, IsAllowed, IsRejected, RejectReason, HoursRatio, " + Environment.NewLine +
                                    "        ESCHours, UsableHours, IsDisOneHour) " + Environment.NewLine +
                                    "VALUES (@DutyIndex, @DepNo, @DepName, @EmpNo, @EmpName, @DutyType, @DutyDateStart, @StartTime, @DutyDateEnd, @EndTime, " + Environment.NewLine +
                                    "        @DutyHours, @BuMan, @BuDate, @AssignReason, @Remark, @Inspector, @IsAllowed, @IsRejected, @RejectReason, " + Environment.NewLine +
                                    "        @HoursRatio, @ESCHours, @UsableHours, @IsDisOneHour)";

                try
                {
                    if (eEmpNo_INS.Text.Trim() != "")
                    {
                        sdsEmpDutyDataDetail.InsertParameters.Clear();
                        sdsEmpDutyDataDetail.InsertCommand = vInsertStr;
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("DutyIndex", DbType.String, vDutyIndex));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo_INS.Text.Trim() != "") ? eDepNo_INS.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("DepName", DbType.String, (eDepName_INS.Text.Trim() != "") ? eDepName_INS.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("EmpNo", DbType.String, (eEmpNo_INS.Text.Trim() != "") ? eEmpNo_INS.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("EmpName", DbType.String, (eEmpName_INS.Text.Trim() != "") ? eEmpName_INS.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("DutyType", DbType.String, (eDutyType_INS.Text.Trim() != "") ? eDutyType_INS.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("DutyDateStart", DbType.Date, (eDutyDateS_INS.Text.Trim() != "") ? vDutyDate_S.ToShortDateString() : String.Empty));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("StartTime", DbType.String, (eStartTime_INS.Text.Trim() != "") ? eStartTime_INS.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("DutyDateEnd", DbType.Date, (eDutyDateE_INS.Text.Trim() != "") ? vDutyDate_E.ToShortDateString() : String.Empty));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("EndTime", DbType.String, (eEndTime_INS.Text.Trim() != "") ? eEndTime_INS.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("DutyHours", DbType.Double, (eDutyHours_INS.Text.Trim() != "") ? eDutyHours_INS.Text.Trim() : "0"));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("BuMan", DbType.String, (eBuMan_INS.Text.Trim() != "") ? eBuMan_INS.Text.Trim() : vLoginID));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, vBuDate.ToShortDateString()));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("AssignReason", DbType.String, (eAssignReason_INS.Text.Trim() != "") ? eAssignReason_INS.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("Remark", DbType.String, (eRemark_INS.Text.Trim() != "") ? eRemark_INS.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("Inspector", DbType.String, (eInspector_INS.Text.Trim() != "") ? eInspector_INS.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("IsAllowed", DbType.Boolean, (cbIsAllowed_INS.Checked) ? "true" : "false"));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("IsRejected", DbType.Boolean, (cbIsRejected_INS.Checked) ? "true" : "false"));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("RejectReason", DbType.String, (eRejectReason_INS.Text.Trim() != "") ? eRejectReason_INS.Text.Trim() : String.Empty));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("HoursRatio", DbType.Int32, (eHoursRatio_INS.Text.Trim() != "") ? eHoursRatio_INS.Text.Trim() : "1"));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("ESCHours", DbType.Double, (eESCHours_INS.Text.Trim() != "") ? eESCHours_INS.Text.Trim() : "0"));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("UsableHours", DbType.Double, (eUsableHours_INS.Text.Trim() != "") ? eUsableHours_INS.Text.Trim() : "0"));
                        sdsEmpDutyDataDetail.InsertParameters.Add(new Parameter("IsDisOneHour", DbType.Boolean, (cbIsDisOneHour_INS.Checked) ? "true" : "false"));
                        sdsEmpDutyDataDetail.Insert();
                    }
                    if (vDutyType == "T02")
                    {
                        //Session["EmpDutyIndex"] = vDutyIndex;
                        CalUsedDuty(vDutyIndex, vDutyDate_S, cbIsAllowed_INS.Checked);
                    }
                    else
                    {
                        //Session["EmpDutyIndex"] = "";
                    }
                    fvEmpDutyDataDetail.ChangeMode(FormViewMode.ReadOnly);
                    this.gridEmpDataList.DataBind();
                    this.gridEmpDutyDataList.DataBind();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
        }

        protected void sdsEmpDutyDataDetail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridEmpDutyDataList.DataBind();
                plSearch.Visible = true;
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('刪除完成！')");
                Response.Write("</" + "Script>");
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + e.Exception.Message + "')");
                Response.Write("</" + "Script>");
            }
            this.gridEmpDataList.DataBind();
            this.gridEmpDutyDataList.DataBind();
        }
    }
}