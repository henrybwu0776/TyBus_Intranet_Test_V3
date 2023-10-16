using Amaterasu_Function;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Threading;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class RentDailyCarCount : Page
    {
        PublicFunction PF = new PublicFunction(); //加入公用程式碼參考
        private string vLoginID = "";
        private string vDepNo = "";
        private string vComputerName = ""; //2021.09.27 新增
        private string vConnStr = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"] != null)
            {
                //顯示民國年
                CultureInfo Cal = new CultureInfo("zh-TW");
                Cal.DateTimeFormat.Calendar = new TaiwanCalendar();
                Cal.DateTimeFormat.YearMonthPattern = "民國 yyy 年 MM 月";
                Thread.CurrentThread.CurrentCulture = Cal;

                vLoginID = (Session["LoginID"] != null) ? Session["LoginID"].ToString().Trim() : "";
                vDepNo = (Session["LoginDepNo"] != null) ? Session["LoginDepNo"].ToString().Trim() : "";
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
                        eCalYear_Search.Text = (DateTime.Today.AddMonths(-1).Year - 1911).ToString();
                        eCalMonth_Search.Text = DateTime.Today.AddMonths(-1).Month.ToString();
                        string vSQLStr = "";
                        if (Int32.Parse(vDepNo) <= 10)
                        {
                            vSQLStr = "select cast(null as varchar) DepNo, cast(null as varchar) DepName " + Environment.NewLine +
                                      "union all " + Environment.NewLine +
                                      "SELECT DEPNO, DEPNO + '-' + NAME AS DepName FROM DEPARTMENT WHERE (ISNULL(InSHReport, 'X') = 'V')";
                        }
                        else
                        {
                            vSQLStr = "SELECT d.DEPNO, d.DEPNO + '-' + d.NAME AS DepName " + Environment.NewLine +
                                      "  FROM DEPARTMENT d " + Environment.NewLine +
                                      " WHERE (ISNULL(InSHReport, 'X') = 'V') " + Environment.NewLine +
                                      "   AND d.DepNo in (SELECT DepNo FROM EmployeeDepNo e where e.EmpNo = '" + vLoginID + "' and isnull(e.IsUsed, 0) = 1)";
                        }
                        sdsDepNo_Search.SelectCommand = vSQLStr;
                        ddlDepNo_Search.DataBind();
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

        /// <summary>
        /// 組合 SQL 查詢字串
        /// </summary>
        /// <returns></returns>
        private string getSelectStr(string fCalDate)
        {
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vResultStr = "select IndexNo, DepNo, (select [Name] from Department where DepNo = a.DepNo) DepName, " + Environment.NewLine +
                                "       CalDate, UseableCount, UseableCount_Red, UseableCount_Green, NeedCarCount " + Environment.NewLine +
                                "  from RentDailyCarCount a " + Environment.NewLine +
                                " where a.CalDate = '" + fCalDate + "' " + Environment.NewLine +
                                vWStr_DepNo +
                                " order by IndexNo";
            return vResultStr;
        }

        /// <summary>
        /// 取回要顯示在行事曆上資料的查詢字串
        /// </summary>
        /// <param name="fCalDate"></param>
        /// <returns></returns>
        private string GetDepCarCount(string fCalDate)
        {
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vResultStr = "select DepNo, (select [Name] from Department where DepNo = a.DepNo) DepName, UseableCount " + Environment.NewLine +
                                "  from RentDailyCarCount a " + Environment.NewLine +
                                " where CalDate = '" + fCalDate + "' " + Environment.NewLine +
                                vWStr_DepNo +
                                " order by DepNo ";
            return vResultStr;
        }

        /// <summary>
        /// 批次填寫單站車輛可用數
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbBatchInsert_Click(object sender, EventArgs e)
        {
            string vSQLStr_Temp = "";
            string vDepNo_Temp = (eDepNo_Search.Text.Trim() != "") ? eDepNo_Search.Text.Trim() : ddlDepNo_Search.SelectedValue.Trim();
            string vCarCount_Temp = (eCarCount_Search.Text.Trim() != "") ? eCarCount_Search.Text.Trim() : "0";
            if (vDepNo_Temp.Trim() != "")
            {
                if (Int32.Parse(vCarCount_Temp) >= 0)
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    DateTime vCalDate_Start = DateTime.Parse(Int32.Parse(eCalYear_Search.Text.Trim()).ToString("D4") + "/" + Int32.Parse(eCalMonth_Search.Text.Trim()).ToString("D2") + "/01");
                    DateTime vCalDate_End = DateTime.Parse(PF.GetMonthLastDay(vCalDate_Start, "C"));
                    int vMonthDays = vCalDate_End.Day;
                    try
                    {
                        for (int i = 0; i < vMonthDays; i++)
                        {
                            DateTime vCalDate_Temp = vCalDate_Start.AddDays(i);
                            string vIndexNo_Temp = vCalDate_Temp.Year.ToString("D4") + vCalDate_Temp.Month.ToString("D2") + vCalDate_Temp.Day.ToString("D2") + Int32.Parse(vDepNo_Temp.Trim()).ToString("D2");
                            vSQLStr_Temp = "select count(IndexNo) RCount from RentDailyCarCount where IndexNo = '" + vIndexNo_Temp + "' ";
                            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "RCount"));
                            using (SqlDataSource dsTemp = new SqlDataSource())
                            {
                                dsTemp.ConnectionString = vConnStr;
                                if (vRCount == 0)
                                {
                                    //新增可用車輛數資料
                                    vSQLStr_Temp = "insert into RentDailyCarCount(IndexNo, DepNo, CalDate, UseableCount, Remark, BuMan, BuDate)" + Environment.NewLine +
                                                   "values (@IndexNo, @DepNo, @CalDate, @UseableCount, @Remark, @BuMan, @BuDate)";
                                    dsTemp.InsertCommand = vSQLStr_Temp;
                                    dsTemp.InsertParameters.Clear();
                                    dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo_Temp));
                                    dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo_Temp));
                                    dsTemp.InsertParameters.Add(new Parameter("CalDate", DbType.Date, vCalDate_Temp.ToShortDateString()));
                                    dsTemp.InsertParameters.Add(new Parameter("UseableCount", DbType.Double, vCarCount_Temp));
                                    dsTemp.InsertParameters.Add(new Parameter("Remark", DbType.String, "批次寫入"));
                                    dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                    dsTemp.InsertParameters.Add(new Parameter("BuDate", DbType.Date, DateTime.Today.ToShortDateString()));
                                    dsTemp.Insert();
                                }
                                else if (vRCount > 0)
                                {
                                    //異動可用車輛數資料
                                    //要先複製一份資料到異動檔
                                    vSQLStr_Temp = "select max(Items) MaxItem from RentDailyCarCount_History where IndexNo = '" + vIndexNo_Temp + "' ";
                                    string vMaxItemStr = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItem");
                                    int vNewItem = (vMaxItemStr != "") ? Int32.Parse(vMaxItemStr) + 1 : 1;
                                    string vIndexNoItem_Temp = vIndexNo_Temp + vNewItem.ToString("D4");
                                    string vItems_Temp = vNewItem.ToString("D4");
                                    string vModifyMode_Temp = "EDIT";
                                    vSQLStr_Temp = "insert into RentDailyCarCount_History" + Environment.NewLine +
                                                   "       (IndexNoItem, Items, ModifyMode, IndexNo, DepNo, CalDate, UseableCount, UseableCount_Red, " + Environment.NewLine +
                                                   "        UseableCount_Green, NeedCarCount, Remark, BuMan, BuDate, ModifyMan, ModifyDate) " + Environment.NewLine +
                                                   "select @IndexNoItem, @Items, @ModifyMode, IndexNo, DepNo, CalDate, UseableCount, UseableCount_Red, " + Environment.NewLine +
                                                   "       UseableCount_Green, NeedCarCount, Remark, BuMan, BuDate, ModifyMan, ModifyDate " + Environment.NewLine +
                                                   "  from RentDailyCarCount " + Environment.NewLine +
                                                   " where IndexNo = @IndexNo";
                                    dsTemp.InsertCommand = vSQLStr_Temp;
                                    dsTemp.InsertParameters.Clear();
                                    dsTemp.InsertParameters.Add(new Parameter("IndexNoItem", DbType.String, vIndexNoItem_Temp.Trim()));
                                    dsTemp.InsertParameters.Add(new Parameter("Items", DbType.String, vItems_Temp.Trim()));
                                    dsTemp.InsertParameters.Add(new Parameter("ModifyMode", DbType.String, vModifyMode_Temp.Trim()));
                                    dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo_Temp));
                                    dsTemp.Insert();
                                    //正式開始寫入異動資料
                                    vSQLStr_Temp = "update RentDailyCarCount " + Environment.NewLine +
                                                   "   set UseableCount = @UseableCount, Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate " + Environment.NewLine +
                                                   " where IndexNo = @IndexNo";
                                    dsTemp.UpdateCommand = vSQLStr_Temp;
                                    dsTemp.UpdateParameters.Clear();
                                    dsTemp.UpdateParameters.Add(new Parameter("UseableCount", DbType.Double, vCarCount_Temp));
                                    dsTemp.UpdateParameters.Add(new Parameter("Remark", DbType.String, "批次寫入"));
                                    dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                    dsTemp.UpdateParameters.Add(new Parameter("ModifyDate", DbType.Date, DateTime.Today.ToShortDateString()));
                                    dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo_Temp));
                                    dsTemp.Update();
                                }
                            }
                        }
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('批次寫入完成！')");
                        Response.Write("</" + "Script>");
                    }
                    catch (Exception eMessage)
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                        Response.Write("</" + "Script>");
                        //throw;
                    }

                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('車輛可用數輸入錯誤 (不可為負數)！')");
                    Response.Write("</" + "Script>");
                    eCarCount_Search.Focus();
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先挑選單位！')");
                Response.Write("</" + "Script>");
                eDepNo_Search.Focus();
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 繫結資料
        /// </summary>
        private void OpenListData()
        {
            string vSQLStr_Temp = "";
            DateTime vSelectDate = calRentDailyCarCount.SelectedDate;
            string vSelDateStr = vSelectDate.Year.ToString() + "/" + vSelectDate.ToString("MM/dd");

            vSQLStr_Temp = getSelectStr(vSelDateStr);
            sdsRentDailyCarCount_List.SelectCommand = vSQLStr_Temp;
            gridRentDailyCarCountList.DataSourceID = "sdsRentDailyCarCount_List";
            gridRentDailyCarCountList.DataBind();
        }

        protected void calRentDailyCarCount_SelectionChanged(object sender, EventArgs e)
        {
            //行事曆日期變更
            OpenListData();
        }

        protected void calRentDailyCarCount_DayRender(object sender, DayRenderEventArgs e)
        {
            if (e.Day.IsOtherMonth)
            {
                e.Cell.ForeColor = System.Drawing.Color.White;
            }
            string vServiceDateStr = e.Day.Date.Year.ToString("D4") + "/" + e.Day.Date.ToString("MM/dd");
            string vSelStr = GetDepCarCount(vServiceDateStr);

            //取回要顯示的資料
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSelStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    ListBox lbTemp = new ListBox();
                    int vIndex = 0;
                    while (drTemp.Read())
                    {
                        lbTemp.Items.Insert(vIndex, drTemp["DepName"].ToString() + " : " + drTemp["UseableCount"].ToString() + " 輛可用");
                        vIndex++;
                    }
                    lbTemp.Rows = vIndex;
                    lbTemp.Attributes.Remove("Style");
                    lbTemp.Attributes.Add("Style", "Width: 95%; Height:120px");

                    e.Cell.Controls.Add(new LiteralControl("<br />"));
                    e.Cell.Controls.Add(lbTemp);
                }
            }
        }

        protected void ddlDepNo_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDepNo_Search.Text = ddlDepNo_Search.SelectedValue.Trim();
        }

        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eIndexNo_Edit = (Label)fvDataDetail.FindControl("eIndexNo_Edit");
            if (eIndexNo_Edit != null)
            {
                try
                {
                    string vIndexNo_Temp = eIndexNo_Edit.Text.Trim();
                    //先複製一份到異動檔
                    string vSQLStr_Temp = "select max(Items) MaxItem from RentDailyCarCount_History where IndexNo = '" + vIndexNo_Temp.Trim() + "' ";
                    string vMaxItem = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItem");
                    int vNewItem = (vMaxItem != "") ? Int32.Parse(vMaxItem) + 1 : 1;
                    string vIndexNoItem = vIndexNo_Temp.Trim() + vNewItem.ToString("D4");
                    string vItems = vNewItem.ToString("D4");
                    vSQLStr_Temp = "insert into RentDailyCarCount_History " + Environment.NewLine +
                                   "       (IndexNoItem, Items, ModifyMode, IndexNo, DepNo, CalDate, UseableCount, UseableCount_Red, UseableCount_Green, " + Environment.NewLine +
                                   "        NeedCarCount, Remark, BuMan, BuDate, ModifyMan, ModifyDate, HisBuMan, HisBuDate) " + Environment.NewLine +
                                   "select '" + vIndexNoItem + "', '" + vItems + "', 'EDIT', IndexNo, DepNo, CalDate, UseableCount, UseableCount_Red, UseableCount_Green, " + Environment.NewLine +
                                   "        NeedCarCount, Remark, BuMan, BuDate, ModifyMan, ModifyDate, '" + vLoginID + "', GetDate() " + Environment.NewLine +
                                   "  from RentDailyCarCount " + Environment.NewLine +
                                   " where IndexNo = '" + vIndexNo_Temp.Trim() + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);

                    //開始更新資料
                    TextBox eUseableCount_Edit = (TextBox)fvDataDetail.FindControl("eUseableCount_Edit");
                    string vUseableCount_Temp = eUseableCount_Edit.Text.Trim();
                    TextBox eRemark_Edit = (TextBox)fvDataDetail.FindControl("eRemark_Edit");
                    string vRemark_Temp = eRemark_Edit.Text.Trim();
                    Label eModifyMan_Edit = (Label)fvDataDetail.FindControl("eModifyMan_Edit");
                    string vModifyMan_Temp = eModifyMan_Edit.Text.Trim();
                    Label eModifyDate_Edit = (Label)fvDataDetail.FindControl("eModifyDate_Edit");
                    DateTime vModifyDate_Temp = (eModifyDate_Edit.Text.Trim() != "") ? DateTime.Parse(eModifyDate_Edit.Text.Trim()) : DateTime.Today;

                    vSQLStr_Temp = "UPDATE RentDailyCarCount " + Environment.NewLine +
                                   "   SET UseableCount = @UseableCount, Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate " + Environment.NewLine +
                                   " WHERE (IndexNo = @IndexNo)";

                    sdsRentDailyCarCount_Detail.UpdateCommand = vSQLStr_Temp;

                    sdsRentDailyCarCount_Detail.UpdateParameters.Clear();
                    sdsRentDailyCarCount_Detail.UpdateParameters.Add(new Parameter("UseableCount", DbType.Double, (vUseableCount_Temp.Trim() != "") ? vUseableCount_Temp.Trim() : String.Empty));
                    sdsRentDailyCarCount_Detail.UpdateParameters.Add(new Parameter("Remark", DbType.String, (vRemark_Temp.Trim() != "") ? vRemark_Temp.Trim() : String.Empty));
                    sdsRentDailyCarCount_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, (vModifyMan_Temp.Trim() != "") ? vModifyMan_Temp.Trim() : vLoginID));
                    sdsRentDailyCarCount_Detail.UpdateParameters.Add(new Parameter("ModifyDate", DbType.Date, (eModifyDate_Edit.Text.Trim() != "") ? vModifyDate_Temp.ToShortDateString() : String.Empty));
                    sdsRentDailyCarCount_Detail.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo_Temp));
                    sdsRentDailyCarCount_Detail.Update();

                    OpenListData();
                    fvDataDetail.ChangeMode(FormViewMode.ReadOnly);
                    fvDataDetail.DataBind();
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

        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eCalDate_INS = (Label)fvDataDetail.FindControl("eCalDate_INS");
            if (eCalDate_INS != null)
            {
                try
                {
                    TextBox eDepNo_INS = (TextBox)fvDataDetail.FindControl("eDepNo_INS");
                    string vDepNo_Temp = eDepNo_INS.Text.Trim();
                    DateTime vCalDate_Temp = DateTime.Parse(eCalDate_INS.Text.Trim());
                    string vIndexNo_Temp = vCalDate_Temp.Year.ToString("D4") + vCalDate_Temp.Month.ToString("D2") + vCalDate_Temp.Day.ToString("D2") + vDepNo_Temp.Trim();
                    TextBox eUseableCount_INS = (TextBox)fvDataDetail.FindControl("eUseableCount");
                    string vUseableCount_Temp = eUseableCount_INS.Text.Trim();
                    TextBox eRemark_INS = (TextBox)fvDataDetail.FindControl("eRemark_INS");
                    string vRemark_Temp = eRemark_INS.Text.Trim();
                    Label eBuMan_INS = (Label)fvDataDetail.FindControl("eBuMan_INS");
                    string vBuMan_Temp = (eBuMan_INS.Text.Trim() != "") ? eBuMan_INS.Text.Trim() : vLoginID;
                    Label eBuDate_INS = (Label)fvDataDetail.FindControl("eBuDate_INS");
                    DateTime vBuDate_Temp = (eBuDate_INS.Text.Trim() != "") ? DateTime.Parse(eBuDate_INS.Text.Trim()) : DateTime.Today;

                    string vSQLStr_Temp = "INSERT INTO RentDailyCarCount " + Environment.NewLine +
                                          "       (IndexNo, DepNo, CalDate, UseableCount, Remark, BuMan, BuDate) " + Environment.NewLine +
                                          "VALUES (@IndexNo, @DepNo, @CalDate, @UseableCount, @Remark, @BuMan, @BuDate)";
                    sdsRentDailyCarCount_Detail.InsertCommand = vSQLStr_Temp;

                    sdsRentDailyCarCount_Detail.InsertParameters.Clear();
                    sdsRentDailyCarCount_Detail.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo_Temp));
                    sdsRentDailyCarCount_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo_Temp));
                    sdsRentDailyCarCount_Detail.InsertParameters.Add(new Parameter("CalDate", DbType.Date, vCalDate_Temp.ToShortDateString()));
                    sdsRentDailyCarCount_Detail.InsertParameters.Add(new Parameter("UseableCount", DbType.Double, (vUseableCount_Temp != "") ? vUseableCount_Temp : String.Empty));
                    sdsRentDailyCarCount_Detail.InsertParameters.Add(new Parameter("Remark", DbType.String, (vRemark_Temp.Trim() != "") ? vRemark_Temp.Trim() : String.Empty));
                    sdsRentDailyCarCount_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vBuMan_Temp.Trim()));
                    sdsRentDailyCarCount_Detail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, (eBuDate_INS.Text.Trim() != "") ? vBuDate_Temp.ToShortDateString() : String.Empty));
                    sdsRentDailyCarCount_Detail.Insert();

                    OpenListData();
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

        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            Label eIndexNo_List = (Label)fvDataDetail.FindControl("eIndexNo_List");
            if (eIndexNo_List != null)
            {
                string vIndexNo_Temp = eIndexNo_List.Text.Trim();
                if (vIndexNo_Temp != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    try
                    {
                        ///先存一份到異動記錄
                        string vSQLStr_Temp = "select max(Items) MaxItem from RentDailyCarCount_History where IndexNo = '" + vIndexNo_Temp.Trim() + "' ";
                        string vMaxItem = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItem");
                        int vNewItem = (vMaxItem != "") ? Int32.Parse(vMaxItem) + 1 : 1;
                        string vIndexNoItem = vIndexNo_Temp.Trim() + vNewItem.ToString("D4");
                        string vItems = vNewItem.ToString("D4");
                        vSQLStr_Temp = "insert into RentDailyCarCount_History " + Environment.NewLine +
                                       "       (IndexNoItem, Items, ModifyMode, IndexNo, DepNo, CalDate, UseableCount, UseableCount_Red, UseableCount_Green, " + Environment.NewLine +
                                       "        NeedCarCount, Remark, BuMan, BuDate, ModifyMan, ModifyDate, HisBuMan, HisBuDate) " + Environment.NewLine +
                                       "select '" + vIndexNoItem + "', '" + vItems + "', 'DEL', IndexNo, DepNo, CalDate, UseableCount, UseableCount_Red, UseableCount_Green, " + Environment.NewLine +
                                       "        NeedCarCount, Remark, BuMan, BuDate, ModifyMan, ModifyDate, '" + vLoginID + "', GetDate() " + Environment.NewLine +
                                       "  from RentDailyCarCount " + Environment.NewLine +
                                       " where IndexNo = '" + vIndexNo_Temp.Trim() + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        //刪除資料
                        vSQLStr_Temp = "delete RentDailyCarCount where IndexNo = @IndexNo";
                        sdsRentDailyCarCount_Detail.DeleteCommand = vSQLStr_Temp;
                        sdsRentDailyCarCount_Detail.DeleteParameters.Clear();
                        sdsRentDailyCarCount_Detail.DeleteParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo_Temp.Trim()));
                        sdsRentDailyCarCount_Detail.Delete();
                        gridRentDailyCarCountList.DataBind();
                    }
                    catch (Exception eMessage)
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                        Response.Write("</" + "Script>");
                        //throw;
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('請先挑選要刪除的資料！')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        protected void fvDataDetail_DataBound(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            switch (fvDataDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    Label eModifyMan_Edit = (Label)fvDataDetail.FindControl("eModifyMan_Edit");
                    Label eModifyManName_Edit = (Label)fvDataDetail.FindControl("eModifyManName_Edit");
                    Label eModifyDate_Edit = (Label)fvDataDetail.FindControl("eModifyDate_Edit");
                    eModifyMan_Edit.Text = vLoginID;
                    eModifyManName_Edit.Text = Session["LoginName"].ToString().Trim();
                    eModifyDate_Edit.Text = DateTime.Today.ToShortDateString();
                    break;
                case FormViewMode.Insert:
                    TextBox eDepNo_INS = (TextBox)fvDataDetail.FindControl("eDepNo_INS");
                    if (Int32.Parse(vDepNo) <= 10)
                    {
                        eDepNo_INS.Enabled = true;
                        eDepNo_INS.Focus();
                    }
                    else
                    {
                        eDepNo_INS.Text = vDepNo.Trim();
                        eDepNo_INS.Enabled = false;
                        TextBox eUseableCount_INS = (TextBox)fvDataDetail.FindControl("eUseableCount_INS");
                        eUseableCount_INS.Focus();
                    }
                    Label eBuMan_INS = (Label)fvDataDetail.FindControl("eBuMan_INS");
                    Label eBuManName_INS = (Label)fvDataDetail.FindControl("eBuManName_INS");
                    eBuMan_INS.Text = vLoginID;
                    eBuManName_INS.Text = Session["LoginName"].ToString().Trim();
                    Label eBuDate_INS = (Label)fvDataDetail.FindControl("eBuDate_INS");
                    eBuDate_INS.Text = DateTime.Today.ToShortDateString();
                    Label eCalDate_INS = (Label)fvDataDetail.FindControl("eCalDate_INS");
                    eCalDate_INS.Text = calRentDailyCarCount.SelectedDate.ToShortDateString();
                    break;
            }
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_INS = (TextBox)fvDataDetail.FindControl("eDepNo_INS");
            Label eDepName_INS = (Label)fvDataDetail.FindControl("eDepName_INS");
            if (eDepNo_INS != null)
            {
                string vDepNo_Temp = eDepNo_INS.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
                string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName_Temp == "")
                {
                    vDepName_Temp = vDepNo_Temp.Trim();
                    vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                    vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo_INS.Text = vDepNo_Temp.Trim();
                eDepName_INS.Text = vDepName_Temp.Trim();
            }
        }
    }
}