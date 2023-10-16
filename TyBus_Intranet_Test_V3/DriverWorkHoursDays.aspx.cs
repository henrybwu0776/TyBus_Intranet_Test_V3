using Amaterasu_Function;
using System;
using System.Data;
using System.Data.SqlClient;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverWorkHoursDays : System.Web.UI.Page
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
                        eCalYear_Search.Text = (DateTime.Today.Year - 1911).ToString("D3");
                        eCalMonth_Search.Text = (DateTime.Today.Month).ToString("D2");
                        ClearData();
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

        private void ClearData()
        {
            fvDriverWorkHoursData.DataSourceID = "";
            plShowData.Visible = false;
            eDriverNo_Search.Text = "";
        }

        protected void eDriverNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDriverNo = eDriverNo_Search.Text.Trim();
            string vDriverName = "";
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vDriverNo + "' and LeaveDay is null and Type = '20' ";
            vDriverName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDriverName == "")
            {
                vDriverName = vDriverNo;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' and LeaveDay is null and Type = '20' ";
                vDriverNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                if (vDriverNo == "")
                {
                    vDriverName = "";
                }
            }
            eDriverNo_Search.Text = vDriverNo;
            eDriverName_Search.Text = vDriverName;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            string vSQLStr = "";
            plShowData.Visible = false;
            DateTime vThisMonth = DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today, "B"));
            DateTime vCalMonthDate = DateTime.Parse((int.Parse(eCalYear_Search.Text.Trim()) + 1911).ToString() + "/" + int.Parse(eCalMonth_Search.Text.Trim()).ToString("D2") + "/01");
            if (vCalMonthDate <= vThisMonth)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vCalYM = (Int32.Parse(eCalYear_Search.Text.Trim()) + 1911).ToString("D4") + "/" + Int32.Parse(eCalMonth_Search.Text.Trim()).ToString("D2");
                string vCalDate = vCalYM + "/01";
                string vDriverNo = eDriverNo_Search.Text.Trim();
                int vMonthDays = PF.GetMonthDays(DateTime.Parse(vCalDate));

                if (vDriverNo != "")
                {
                    if (vLoginDepNo != "09")
                    {
                        vSQLStr = "select count(EmpNo) RCount from Employee " + Environment.NewLine +
                                  " where EmpNo = '" + vDriverNo + "' and DepNo = '" + vLoginDepNo + "' ";
                    }
                    else
                    {
                        vSQLStr = "select count(EmpNo) RCount from Employee " + Environment.NewLine +
                                  " where EmpNo = '" + vDriverNo + "' ";
                    }
                    int vRCount = int.Parse(PF.GetValue(vConnStr, vSQLStr, "RCount"));

                    if (vRCount > 0)
                    {
                        DataTable dtTemp = new DataTable("DriverWorkHours");
                        dtTemp.Columns.Add("Driver", typeof(String));
                        dtTemp.Columns.Add("DriverName", typeof(String));
                        dtTemp.Columns.Add("RentNumber", typeof(Double));
                        dtTemp.Columns.Add("TotalHR", typeof(String));
                        dtTemp.Columns.Add("Hour01", typeof(String));
                        dtTemp.Columns.Add("WorkState01", typeof(String));
                        dtTemp.Columns.Add("Hour02", typeof(String));
                        dtTemp.Columns.Add("WorkState02", typeof(String));
                        dtTemp.Columns.Add("Hour03", typeof(String));
                        dtTemp.Columns.Add("WorkState03", typeof(String));
                        dtTemp.Columns.Add("Hour04", typeof(String));
                        dtTemp.Columns.Add("WorkState04", typeof(String));
                        dtTemp.Columns.Add("Hour05", typeof(String));
                        dtTemp.Columns.Add("WorkState05", typeof(String));
                        dtTemp.Columns.Add("Hour06", typeof(String));
                        dtTemp.Columns.Add("WorkState06", typeof(String));
                        dtTemp.Columns.Add("Hour07", typeof(String));
                        dtTemp.Columns.Add("WorkState07", typeof(String));
                        dtTemp.Columns.Add("Hour08", typeof(String));
                        dtTemp.Columns.Add("WorkState08", typeof(String));
                        dtTemp.Columns.Add("Hour09", typeof(String));
                        dtTemp.Columns.Add("WorkState09", typeof(String));
                        dtTemp.Columns.Add("Hour10", typeof(String));
                        dtTemp.Columns.Add("WorkState10", typeof(String));
                        dtTemp.Columns.Add("Hour11", typeof(String));
                        dtTemp.Columns.Add("WorkState11", typeof(String));
                        dtTemp.Columns.Add("Hour12", typeof(String));
                        dtTemp.Columns.Add("WorkState12", typeof(String));
                        dtTemp.Columns.Add("Hour13", typeof(String));
                        dtTemp.Columns.Add("WorkState13", typeof(String));
                        dtTemp.Columns.Add("Hour14", typeof(String));
                        dtTemp.Columns.Add("WorkState14", typeof(String));
                        dtTemp.Columns.Add("Hour15", typeof(String));
                        dtTemp.Columns.Add("WorkState15", typeof(String));
                        dtTemp.Columns.Add("Hour16", typeof(String));
                        dtTemp.Columns.Add("WorkState16", typeof(String));
                        dtTemp.Columns.Add("Hour17", typeof(String));
                        dtTemp.Columns.Add("WorkState17", typeof(String));
                        dtTemp.Columns.Add("Hour18", typeof(String));
                        dtTemp.Columns.Add("WorkState18", typeof(String));
                        dtTemp.Columns.Add("Hour19", typeof(String));
                        dtTemp.Columns.Add("WorkState19", typeof(String));
                        dtTemp.Columns.Add("Hour20", typeof(String));
                        dtTemp.Columns.Add("WorkState20", typeof(String));
                        dtTemp.Columns.Add("Hour21", typeof(String));
                        dtTemp.Columns.Add("WorkState21", typeof(String));
                        dtTemp.Columns.Add("Hour22", typeof(String));
                        dtTemp.Columns.Add("WorkState22", typeof(String));
                        dtTemp.Columns.Add("Hour23", typeof(String));
                        dtTemp.Columns.Add("WorkState23", typeof(String));
                        dtTemp.Columns.Add("Hour24", typeof(String));
                        dtTemp.Columns.Add("WorkState24", typeof(String));
                        dtTemp.Columns.Add("Hour25", typeof(String));
                        dtTemp.Columns.Add("WorkState25", typeof(String));
                        dtTemp.Columns.Add("Hour26", typeof(String));
                        dtTemp.Columns.Add("WorkState26", typeof(String));
                        dtTemp.Columns.Add("Hour27", typeof(String));
                        dtTemp.Columns.Add("WorkState27", typeof(String));
                        dtTemp.Columns.Add("Hour28", typeof(String));
                        dtTemp.Columns.Add("WorkState28", typeof(String));
                        dtTemp.Columns.Add("Hour29", typeof(String));
                        dtTemp.Columns.Add("WorkState29", typeof(String));
                        dtTemp.Columns.Add("Hour30", typeof(String));
                        dtTemp.Columns.Add("WorkState30", typeof(String));
                        dtTemp.Columns.Add("Hour31", typeof(String));
                        dtTemp.Columns.Add("WorkState31", typeof(String));

                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            int vTotalHours = 0;
                            int vTotalMins = 0;
                            double vRentNumber = 0;
                            string vHourFieldName = "";
                            string vStateFieldName = "";

                            vSQLStr = "select datepart(day, a.budate) as WorkDays, a.Driver, (select [Name] from Employee where EmpNo = a.Driver) DriverName, " + Environment.NewLine +
                                      "       isnull(a.RentNumber, 0) RentNumber, a.WorkHR, a.WorkMin, " + Environment.NewLine +
                                      "       left((select ClassTxt from DBDICB where ClassNo = a.WorkState and FKey = '行車記錄單A     runsheeta       WORKSTATE'), 1) as WorkState_C " + Environment.NewLine +
                                      "  from RunSheetA a " + Environment.NewLine +
                                      " where convert(varchar(7), a.budate, 111) = '" + vCalYM + "' " + Environment.NewLine +
                                      "   and a.Driver = '" + vDriverNo + "' " + Environment.NewLine +
                                      " order by a.BuDate ";
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            DataRow rowTemp = dtTemp.NewRow();
                            while (drTemp.Read())
                            {
                                rowTemp["Driver"] = Int32.Parse(drTemp["Driver"].ToString()).ToString("D6");
                                rowTemp["DriverName"] = drTemp["DriverName"].ToString();
                                vHourFieldName = "Hour" + Int32.Parse(drTemp["WorkDays"].ToString()).ToString("D2");
                                rowTemp[vHourFieldName] = (Int32.Parse(drTemp["WorkHR"].ToString()) + Int32.Parse(drTemp["WorkMin"].ToString()) > 0) ?
                                                          Int32.Parse(drTemp["WorkHR"].ToString()).ToString("D2") + ":" + Int32.Parse(drTemp["WorkMin"].ToString()).ToString("D2") :
                                                          "X";
                                vStateFieldName = "WorkState" + Int32.Parse(drTemp["WorkDays"].ToString()).ToString("D2");
                                rowTemp[vStateFieldName] = (drTemp["WorkState_C"].ToString() != "平") ? drTemp["WorkState_C"].ToString() : "";
                                vRentNumber = vRentNumber + double.Parse(drTemp["RentNumber"].ToString());
                                vTotalHours = vTotalHours + Int32.Parse(drTemp["WorkHR"].ToString());
                                vTotalMins = vTotalMins + Int32.Parse(drTemp["WorkMin"].ToString());
                            }
                            rowTemp["RentNumber"] = vRentNumber.ToString();
                            vTotalHours = vTotalHours + (vTotalMins / 60);
                            vTotalMins = vTotalMins % 60;
                            rowTemp["TotalHR"] = (vTotalHours + vTotalMins > 0) ? vTotalHours.ToString() + ":" + vTotalMins.ToString("D2") : "X";
                            for (int i = 1; i <= vMonthDays; i++)
                            {
                                vHourFieldName = "Hour" + i.ToString("D2");
                                rowTemp[vHourFieldName] = (rowTemp[vHourFieldName].ToString().Trim() == "") ? "X" : rowTemp[vHourFieldName].ToString().Trim();
                            }
                            dtTemp.Rows.Add(rowTemp);

                            fvDriverWorkHoursData.DataSource = dtTemp;
                            fvDriverWorkHoursData.DataBind();
                            plShowData.Visible = true;

                            string vRecordDriverStr = (eDriverNo_Search.Text.Trim() != "") ? eDriverNo_Search.Text.Trim() : "全部";
                            string vRecordNote = "查詢資料_駕駛員時數統計查詢" + Environment.NewLine +
                                                 "DriverWorkHoursDays.aspx" + Environment.NewLine +
                                                 "行車日期：" + eCalYear_Search.Text.Trim() + "年" + eCalMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                                 "駕駛員：" + vRecordDriverStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                        }
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('在單位 [ " + Session["DepName"].ToString().Trim() + " ] 找不到駕駛員 [ " + vDriverNo + " ] 的資料')");
                        Response.Write("</" + "Script>");
                        ClearData();
                        eCalYear_Search.Focus();
                    }
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('不開放查詢未來的資料！')");
                Response.Write("</" + "Script>");
                ClearData();
                eCalYear_Search.Focus();
            }
        }

        protected void bbClear_Click(object sender, EventArgs e)
        {
            ClearData();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}