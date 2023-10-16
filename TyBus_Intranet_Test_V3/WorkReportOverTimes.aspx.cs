using Amaterasu_Function;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class WorkReportOverTimes : System.Web.UI.Page
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
        private DataTable dtOverDutyList = new DataTable("OverDutyList");

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
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        eCalYear.Text = (DateTime.Today.AddMonths(-1).Year - 1911).ToString("D3");
                        eCalMonth.Text = DateTime.Today.AddMonths(-1).Month.ToString("D2");
                        eSplitMode.Text = rbSplitMode.SelectedValue.Trim();
                        plShowData_Main0.Visible = false;
                        plShowData_Main1.Visible = false;
                        bbExcel.Enabled = false;
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

        private string GetSelStr_Driver(string fMode)
        {
            string vCalYear = (Int32.Parse(eCalYear.Text.Trim()) < 1911) ? (Int32.Parse(eCalYear.Text.Trim()) + 1911).ToString("D4") :
                              (Int32.Parse(eCalYear.Text.Trim()) > 3822) ? (Int32.Parse(eCalYear.Text.Trim()) - 1911).ToString("D4") :
                              eCalYear.Text.Trim();
            DateTime vStartDate = DateTime.Parse(vCalYear + "/" + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2") + "/01");
            DateTime vEndDate = DateTime.Parse(PF.GetMonthLastDay(vStartDate, "B"));
            string vWStr_DateRange = " between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "' ";
            string vSelStr = "";
            switch (fMode)
            {
                case "0": //不區分性別資料時的查詢字串
                    vSelStr = "select a.DepNo, '單位：' + (select [Name] from Department where DepNo = a.DepNo) DepName, " + Environment.NewLine +
                              "       cast(sum(a.TotalMin) as int) / 60 TotalHR, cast(sum(a.TotalMin) as int) % 60 TotalMin " + Environment.NewLine +
                              "  from (" + Environment.NewLine +
                              "        select DepNo, ((WorkHR * 60 + WorkMin) - 480) TotalMin " + Environment.NewLine +
                              "          from RunSheetA " + Environment.NewLine +
                              "         where Budate " + vWStr_DateRange + Environment.NewLine +
                              "           and (WorkHR * 60 + WorkMin) > 480 " + Environment.NewLine +
                              "           and WorkState = '0' " + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select DepNo, (WorkHR * 60 + WorkMin) TotalMin " + Environment.NewLine +
                              "          from RunSheetA " + Environment.NewLine +
                              "         where Budate " + vWStr_DateRange + Environment.NewLine +
                              "           and WorkState <> '0' " + Environment.NewLine +
                              ") a " + Environment.NewLine +
                              " group by a.DepNo " + Environment.NewLine +
                              " order by a.DepNo";
                    break;
                case "1": //要區分性別資料時的查詢字串
                    vSelStr = "select a.DepNo, '單位：' + (select [Name] from Department where DepNo = a.DepNo) DepName, " + Environment.NewLine +
                              "       cast(sum(a.TotalMin_M) as int) / 60 TotalHR_M, cast(sum(a.TotalMin_M) as int) % 60 TotalMin_M, " + Environment.NewLine +
                              "       cast(sum(a.TotalMin_F) as int) / 60 TotalHR_F, cast(sum(a.TotalMin_F) as int) % 60 TotalMin_F " + Environment.NewLine +
                              "  from ( " + Environment.NewLine +
                              "        select DepNo, ((WorkHR * 60 + WorkMin) - 480) TotalMin_M, cast(0 as int) TotalMin_F " + Environment.NewLine +
                              "          from RunSheetA ra " + Environment.NewLine +
                              "         where Budate " + vWStr_DateRange + Environment.NewLine +
                              "           and (WorkHR * 60 + WorkMin) > 480 " + Environment.NewLine +
                              "           and WorkState = '0' " + Environment.NewLine +
                              "           and (select Gender from Employee where EmpNo = ra.Driver) = '男' " + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select DepNo, (WorkHR * 60 + WorkMin) TotalMin_M, cast(0 as int) TotalMin_F " + Environment.NewLine +
                              "          from RunSheetA ra " + Environment.NewLine +
                              "         where Budate " + vWStr_DateRange + Environment.NewLine +
                              "           and WorkState <> '0' " + Environment.NewLine +
                              "           and (select Gender from Employee where EmpNo = ra.Driver) = '男' " + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select DepNo, cast(0 as int) TotalMin_M, ((WorkHR * 60 + WorkMin) - 480) TotalMin_F " + Environment.NewLine +
                              "          from RunSheetA ra " + Environment.NewLine +
                              "         where Budate " + vWStr_DateRange + Environment.NewLine +
                              "           and (WorkHR * 60 + WorkMin) > 480 " + Environment.NewLine +
                              "           and WorkState = '0' " + Environment.NewLine +
                              "           and (select Gender from Employee where EmpNo = ra.Driver) = '女' " + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select DepNo, cast(0 as int) TotalMin_M, (WorkHR * 60 + WorkMin) TotalMin_F " + Environment.NewLine +
                              "          from RunSheetA ra " + Environment.NewLine +
                              "         where Budate " + vWStr_DateRange + Environment.NewLine +
                              "           and WorkState <> '0' " + Environment.NewLine +
                              "           and (select Gender from Employee where EmpNo = ra.Driver) = '女' " + Environment.NewLine +
                              ") a " + Environment.NewLine +
                              " group by a.DepNo " + Environment.NewLine +
                              " order by a.DepNo";
                    break;
            }
            return vSelStr;
        }

        private void GetData_Emp(string fMode)
        {
            string vCalYear = (Int32.Parse(eCalYear.Text.Trim()) < 1911) ? (Int32.Parse(eCalYear.Text.Trim()) + 1911).ToString("D4") :
                              (Int32.Parse(eCalYear.Text.Trim()) > 3822) ? (Int32.Parse(eCalYear.Text.Trim()) - 1911).ToString("D4") :
                              eCalYear.Text.Trim();
            DateTime vStartDate = DateTime.Parse(vCalYear + "/" + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2") + "/01");
            DateTime vEndDate = DateTime.Parse(PF.GetMonthLastDay(vStartDate, "B"));
            string vWStr_DateRange = " between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "' ";
            //先取回指定年月加班單有出的職稱代號和員工類別
            string vDepNo = "";
            string vColName = "";
            string vSelStr = "";
            string vStrTital = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                if (fMode == "0")
                {
                    vSelStr = "select EmpType_C " + Environment.NewLine +
                              "  from ( " + Environment.NewLine +
                              "        select distinct (select[Type] from Employee where EmpNo = a.ApplyMan) EmpType, " + Environment.NewLine +
                              "               (select Title from Employee where EmpNo = a.ApplyMan) + " + Environment.NewLine +
                              "               left((select ClassTxt from DBDICB where ClassNo = (select[Type] from Employee where EmpNo = a.ApplyMan) and FKey = '人事資料檔      EMPLOYEE        type'), 2) EmpType_C " + Environment.NewLine +
                              "          from OverDuty a " + Environment.NewLine +
                              "         where ApplyType <> '03' " + Environment.NewLine +
                              "           and RealDay " + vWStr_DateRange + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select distinct (select[Type] from Employee where EmpNo = a.ApplyMan) EmpType, " + Environment.NewLine +
                              "               left((select ClassTxt from DBDICB where ClassNo = (select[Type] from Employee where EmpNo = a.ApplyMan) and FKey = '人事資料檔      EMPLOYEE        type'), 2) + '合計' EmpType_C " + Environment.NewLine +
                              "          from OverDuty a " + Environment.NewLine +
                              "         where ApplyType <> '03' " + Environment.NewLine +
                              "           and RealDay " + vWStr_DateRange + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select '99' as EmpType, '總計' as EmpType_C " + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select '' as EmpType, '單位' as EmpType_C " + Environment.NewLine +
                              ") z order by z.EmpType, EmpType_C";
                }
                else if (fMode == "1")
                {
                    vSelStr = "select EmpType_C " + Environment.NewLine +
                              "  from ( " + Environment.NewLine +
                              "        select distinct (select[Type] +'0' from Employee where EmpNo = a.ApplyMan) EmpType, " + Environment.NewLine +
                              "               (select Title from Employee where EmpNo = a.ApplyMan) + " + Environment.NewLine +
                              "               left((select ClassTxt from DBDICB where ClassNo = (select[Type] from Employee where EmpNo = a.ApplyMan) and FKey = '人事資料檔      EMPLOYEE        type'), 2) + " + Environment.NewLine +
                              "               ' (男)' EmpType_C " + Environment.NewLine +
                              "          from OverDuty a " + Environment.NewLine +
                              "         where ApplyType <> '03' " + Environment.NewLine +
                              "           and RealDay " + vWStr_DateRange + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select distinct (select[Type] +'0' from Employee where EmpNo = a.ApplyMan) EmpType, " + Environment.NewLine +
                              "               (select Title from Employee where EmpNo = a.ApplyMan) + " + Environment.NewLine +
                              "               left((select ClassTxt from DBDICB where ClassNo = (select[Type] from Employee where EmpNo = a.ApplyMan) and FKey = '人事資料檔      EMPLOYEE        type'), 2) + " + Environment.NewLine +
                              "               ' (女)' EmpType_C " + Environment.NewLine +
                              "          from OverDuty a " + Environment.NewLine +
                              "         where ApplyType <> '03' " + Environment.NewLine +
                              "           and RealDay " + vWStr_DateRange + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select distinct (select[Type] +'1' from Employee where EmpNo = a.ApplyMan) EmpType, " + Environment.NewLine +
                              "               left((select ClassTxt from DBDICB where ClassNo = (select[Type] from Employee where EmpNo = a.ApplyMan) and FKey = '人事資料檔      EMPLOYEE        type'), 2) +' (男)合計' EmpType_C " + Environment.NewLine +
                              "          from OverDuty a " + Environment.NewLine +
                              "         where ApplyType <> '03' " + Environment.NewLine +
                              "           and RealDay " + vWStr_DateRange + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select distinct (select[Type] +'1' from Employee where EmpNo = a.ApplyMan) EmpType, " + Environment.NewLine +
                              "               left((select ClassTxt from DBDICB where ClassNo = (select[Type] from Employee where EmpNo = a.ApplyMan) and FKey = '人事資料檔      EMPLOYEE        type'), 2) +' (女)合計' EmpType_C " + Environment.NewLine +
                              "          from OverDuty a " + Environment.NewLine +
                              "         where ApplyType <> '03' " + Environment.NewLine +
                              "           and RealDay " + vWStr_DateRange + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select distinct (select[Type] +'2' from Employee where EmpNo = a.ApplyMan) EmpType, " + Environment.NewLine +
                              "               left((select ClassTxt from DBDICB where ClassNo = (select[Type] from Employee where EmpNo = a.ApplyMan) and FKey = '人事資料檔      EMPLOYEE        type'), 2) +'合計' EmpType_C " + Environment.NewLine +
                              "          from OverDuty a " + Environment.NewLine +
                              "         where ApplyType <> '03' " + Environment.NewLine +
                              "           and RealDay " + vWStr_DateRange + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select '990' as EmpType, '總計 (男)' as EmpType_C " + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select '990' as EmpType, '總計 (女)' as EmpType_C " + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select '991' as EmpType, '總計' as EmpType_C " + Environment.NewLine +
                              "         union all " + Environment.NewLine +
                              "        select '' as EmpType, '單位' as EmpType_C " + Environment.NewLine +
                              ") z order by z.EmpType, EmpType_C";
                }
                SqlCommand cmdTemp = new SqlCommand(vSelStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    //先清除 DataTable 的內容
                    dtOverDutyList.Clear();
                    while (drTemp.Read())
                    {
                        vColName = drTemp["EmpType_C"].ToString().Trim();
                        if (vColName == "單位")
                        {
                            dtOverDutyList.Columns.Add(vColName, typeof(String));
                        }
                        else
                        {
                            dtOverDutyList.Columns.Add(vColName, typeof(Double));
                            vStrTital = (vStrTital == "") ? "[" + vColName + "]" : vStrTital + ", [" + vColName + "]";
                        }
                    }
                }
            }
            using (SqlConnection connDepNo = new SqlConnection(vConnStr))
            {
                string vGetDataStr = "";
                string vSelDepNoStr = "select distinct DepNo " + Environment.NewLine +
                                      "  from OverDuty " + Environment.NewLine +
                                      " where ApplyType <> '03' " + Environment.NewLine +
                                      "   and RealDay " + vWStr_DateRange + Environment.NewLine +
                                      " order by DepNo";
                SqlCommand cmdDepNo = new SqlCommand(vSelDepNoStr, connDepNo);
                connDepNo.Open();
                SqlDataReader drDepNo = cmdDepNo.ExecuteReader();
                if (drDepNo.HasRows)
                {
                    while (drDepNo.Read())
                    {
                        vDepNo = drDepNo["DepNo"].ToString().Trim();
                        if (fMode == "0")
                        {
                            vGetDataStr = "select * from " + Environment.NewLine +
                                          "   ( " + Environment.NewLine +
                                          "    select [Hours], (select Title from Employee where EmpNo = a.ApplyMan) + " + Environment.NewLine +
                                          "           left((select ClassTxt from DBDICB where ClassNo = (select[Type] from Employee where EmpNo = a.ApplyMan) and FKey = '人事資料檔      EMPLOYEE        type'), 2) EmpType_C " + Environment.NewLine +
                                          "      from OverDuty a " + Environment.NewLine +
                                          "     where ApplyType<> '03' " + Environment.NewLine +
                                          "       and RealDay " + vWStr_DateRange + Environment.NewLine +
                                          "       and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                          "     union all " + Environment.NewLine +
                                          "     select [Hours], left((select ClassTxt from DBDICB where ClassNo = (select[Type] from Employee where EmpNo = a.ApplyMan) and FKey = '人事資料檔      EMPLOYEE        type'), 2) +'合計' EmpType_C " + Environment.NewLine +
                                          "       from OverDuty a " + Environment.NewLine +
                                          "      where ApplyType<> '03' " + Environment.NewLine +
                                          "        and RealDay " + vWStr_DateRange + Environment.NewLine +
                                          "        and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                          "      union all " + Environment.NewLine +
                                          "     select [Hours], '總計' as EmpType_C " + Environment.NewLine +
                                          "       from OverDuty a " + Environment.NewLine +
                                          "      where ApplyType <> '03' " + Environment.NewLine +
                                          "        and RealDay " + vWStr_DateRange + Environment.NewLine +
                                          "        and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                          "   ) t1" + Environment.NewLine +
                                          "pivot " + Environment.NewLine +
                                          "(sum([Hours]) for EmpType_C in (" + vStrTital + ")) t2";
                        }
                        else if (fMode == "1")
                        {
                            vGetDataStr = "select * from " + Environment.NewLine +
                                          "   ( " + Environment.NewLine +
                                          "    select [Hours], (select Title from Employee where EmpNo = a.ApplyMan) + " + Environment.NewLine +
                                          "           left((select ClassTxt from DBDICB where ClassNo = (select[Type] from Employee where EmpNo = a.ApplyMan) and FKey = '人事資料檔      EMPLOYEE        type'), 2) + " + Environment.NewLine +
                                          "           ' (' + (select Gender from Employee where EmpNo = a.ApplyMan) +')' EmpType_C " + Environment.NewLine +
                                          "      from OverDuty a " + Environment.NewLine +
                                          "     where ApplyType<> '03' " + Environment.NewLine +
                                          "       and RealDay " + vWStr_DateRange + Environment.NewLine +
                                          "       and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                          "     union all " + Environment.NewLine +
                                          "    select [Hours], left((select ClassTxt from DBDICB where ClassNo = (select[Type] from Employee where EmpNo = a.ApplyMan) and FKey = '人事資料檔      EMPLOYEE        type'), 2) + " + Environment.NewLine +
                                          "           ' (' + (select Gender from Employee where EmpNo = a.ApplyMan) +')合計' EmpType_C " + Environment.NewLine +
                                          "       from OverDuty a " + Environment.NewLine +
                                          "      where ApplyType<> '03' " + Environment.NewLine +
                                          "        and RealDay " + vWStr_DateRange + Environment.NewLine +
                                          "        and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                          "      union all " + Environment.NewLine +
                                          "     select [Hours], left((select ClassTxt from DBDICB where ClassNo = (select[Type] from Employee where EmpNo = a.ApplyMan) and FKey = '人事資料檔      EMPLOYEE        type'), 2) +'合計' EmpType_C " + Environment.NewLine +
                                          "       from OverDuty a " + Environment.NewLine +
                                          "      where ApplyType<> '03' " + Environment.NewLine +
                                          "        and RealDay " + vWStr_DateRange + Environment.NewLine +
                                          "        and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                          "      union all " + Environment.NewLine +
                                          "     select [Hours], '總計 (' + (select Gender from Employee where EmpNo = a.ApplyMan) +')' as EmpType_C " + Environment.NewLine +
                                          "       from OverDuty a " + Environment.NewLine +
                                          "      where ApplyType <> '03' " + Environment.NewLine +
                                          "        and RealDay " + vWStr_DateRange + Environment.NewLine +
                                          "        and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                          "      union all " + Environment.NewLine +
                                          "     select [Hours], '總計' as EmpType_C " + Environment.NewLine +
                                          "       from OverDuty a " + Environment.NewLine +
                                          "      where ApplyType <> '03' " + Environment.NewLine +
                                          "        and RealDay " + vWStr_DateRange + Environment.NewLine +
                                          "        and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                          "   ) t1" + Environment.NewLine +
                                          "pivot " + Environment.NewLine +
                                          "(sum([Hours]) for EmpType_C in (" + vStrTital + ")) t2";
                        }
                        using (SqlConnection connGetData = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdGetData = new SqlCommand(vGetDataStr, connGetData);
                            connGetData.Open();
                            SqlDataReader drGetData = cmdGetData.ExecuteReader();
                            if (drGetData.HasRows)
                            {
                                while (drGetData.Read())
                                {
                                    DataRow rowsTemp = dtOverDutyList.NewRow();
                                    for (int i = 0; i < dtOverDutyList.Columns.Count; i++)
                                    {
                                        vColName = dtOverDutyList.Columns[i].ColumnName.Trim();
                                        rowsTemp[i] = (vColName == "單位") ? vDepNo : (Convert.IsDBNull(drGetData[vColName])) ? "0" : drGetData[vColName].ToString().Trim();
                                    }
                                    dtOverDutyList.Rows.Add(rowsTemp);
                                }
                            }
                            drGetData.Close();
                            cmdGetData.Cancel();
                        }
                    }
                }
                drDepNo.Close();
                cmdDepNo.Cancel();
            }
        }

        private void SearchData()
        {
            string vMode = eSplitMode.Text.Trim();
            string vSelectStr_Driver = GetSelStr_Driver(vMode);
            switch (vMode)
            {
                case "0":
                    //駕駛員時數部份
                    sdsShowData_Driver0.SelectCommand = "";
                    sdsShowData_Driver1.SelectCommand = "";
                    sdsShowData_Driver0.SelectCommand = vSelectStr_Driver;
                    gridShowData_Driver0.DataBind();

                    //內勤人員加班
                    GetData_Emp(vMode);
                    gridShowData_Emp0.DataSource = dtOverDutyList;
                    gridShowData_Emp0.DataBind();

                    //顯示結果
                    plShowData_Main0.Visible = true;
                    plShowData_Main1.Visible = false;
                    bbExcel.Enabled = ((gridShowData_Emp0.Rows.Count > 0) && (gridShowData_Driver0.Rows.Count > 0));
                    break;
                case "1":
                    //駕駛員時數部份
                    sdsShowData_Driver0.SelectCommand = "";
                    sdsShowData_Driver1.SelectCommand = "";
                    sdsShowData_Driver1.SelectCommand = vSelectStr_Driver;
                    gridShowData_Driver1.DataBind();

                    //內勤人員加班
                    GetData_Emp(vMode);
                    gridShowData_Emp1.DataSource = dtOverDutyList;
                    gridShowData_Emp1.DataBind();

                    //顯示結果
                    plShowData_Main0.Visible = false;
                    plShowData_Main1.Visible = true;
                    bbExcel.Enabled = ((gridShowData_Emp1.Rows.Count > 0) && (gridShowData_Driver1.Rows.Count > 0));
                    break;
            }
        }

        protected void rbSplitMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            eSplitMode.Text = rbSplitMode.SelectedValue.Trim();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            SearchData();
            string vCalYMStr = eCalYear.Text.Trim() + " 年 " + eCalMonth.Text.Trim() + " 月";
            string vSplitModeStr = rbSplitMode.SelectedItem.Text.Trim();
            string vRecordNote = "查詢資料_工作會報用加班時數統計" + Environment.NewLine +
                                 "WorkReportOverTimes.aspx" + Environment.NewLine +
                                 "計算年月：" + vCalYMStr + Environment.NewLine +
                                 "區分性別：" + vSplitModeStr;
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vFileName = eCalYear.Text.Trim() + "年" + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2") + "月加班時數_工作會報用";
            string vMode = eSplitMode.Text.Trim();
            //準備匯出EXCEL
            XSSFWorkbook wbExcel = new XSSFWorkbook();
            //Excel 工作表
            XSSFSheet wsExcel;

            //設定標題欄位的格式
            XSSFCellStyle csTitle = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中

            //設定標題欄位的字體格式
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

            XSSFCellStyle csData_Red = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            XSSFFont fontData_Red = (XSSFFont)wbExcel.CreateFont();
            fontData_Red.FontHeightInPoints = 12; //字體大小
            fontData_Red.Color = IndexedColors.Red.Index; //字體顏色
            csData_Red.SetFont(fontData_Red);

            XSSFCellStyle csData_Int = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            XSSFDataFormat format = (XSSFDataFormat)wbExcel.CreateDataFormat();
            csData_Int.DataFormat = format.GetFormat("###,##0");

            string vHeaderText = "";
            string vSheetName = "駕駛員統計";
            string vCellName = "";
            string vCellName_S = "";
            string vCellName_E = "";
            int vLinesNo = 0;
            int vStrCode_1 = 0;
            int vStrCode_2 = 0;
            //新增一個工作表
            wsExcel = (XSSFSheet)wbExcel.CreateSheet(vSheetName);
            //寫入標題列
            vLinesNo = 0;
            wsExcel.CreateRow(vLinesNo);
            if (vMode == "0")
            {
                for (int i = 0; i < gridShowData_Driver0.Columns.Count; i++)
                {
                    vHeaderText = gridShowData_Driver0.Columns[i].HeaderText.Trim();
                    wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                }
                for (int j = 0; j < gridShowData_Driver0.Rows.Count; j++)
                {
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    for (int k = 0; k < gridShowData_Driver0.Columns.Count; k++)
                    {
                        vHeaderText = gridShowData_Driver0.Columns[k].HeaderText.Trim();
                        if ((vHeaderText == "單位") || (vHeaderText == "列標籤"))
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(k);
                            wsExcel.GetRow(vLinesNo).GetCell(k).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(k).SetCellValue(gridShowData_Driver0.Rows[j].Cells[k].Text.Trim());
                            wsExcel.GetRow(vLinesNo).GetCell(k).CellStyle = csData;
                        }
                        else
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(k);
                            wsExcel.GetRow(vLinesNo).GetCell(k).SetCellType(CellType.Numeric);
                            wsExcel.GetRow(vLinesNo).GetCell(k).SetCellValue(Double.Parse(gridShowData_Driver0.Rows[j].Cells[k].Text.Trim()));
                            wsExcel.GetRow(vLinesNo).GetCell(k).CellStyle = csData;
                        }
                    }
                }
                vLinesNo++;
                wsExcel.CreateRow(vLinesNo);
                for (int i = 0; i < gridShowData_Driver0.Columns.Count; i++)
                {
                    vHeaderText = gridShowData_Driver0.Columns[i].HeaderText.Trim();
                    if (vHeaderText == "單位")
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(i);
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue("");
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    else if (vHeaderText == "列標籤")
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(i);
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue("總計");
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    else
                    {
                        vStrCode_1 = (i % 26) + 65;
                        vStrCode_2 = (i / 26);
                        vCellName = (vStrCode_2 != 0) ? ((char)(vStrCode_2 + 64)).ToString() + ((char)vStrCode_1).ToString() : ((char)vStrCode_1).ToString();

                        vCellName_S = vCellName + "2";
                        vCellName_E = vCellName + vLinesNo.ToString();
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("sum(" + vCellName_S + ":" + vCellName_E + ")");
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                }
            }
            else if (vMode == "1")
            {
                for (int i = 0; i < gridShowData_Driver1.Columns.Count; i++)
                {
                    vHeaderText = gridShowData_Driver1.Columns[i].HeaderText.Trim();
                    wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                }
                for (int j = 0; j < gridShowData_Driver1.Rows.Count; j++)
                {
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    for (int k = 0; k < gridShowData_Driver1.Columns.Count; k++)
                    {
                        vHeaderText = gridShowData_Driver1.Columns[k].HeaderText.Trim();
                        if ((vHeaderText == "單位") || (vHeaderText == "列標籤"))
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(k);
                            wsExcel.GetRow(vLinesNo).GetCell(k).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(k).SetCellValue(gridShowData_Driver1.Rows[j].Cells[k].Text.Trim());
                            wsExcel.GetRow(vLinesNo).GetCell(k).CellStyle = csData;
                        }
                        else
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(k);
                            wsExcel.GetRow(vLinesNo).GetCell(k).SetCellType(CellType.Numeric);
                            wsExcel.GetRow(vLinesNo).GetCell(k).SetCellValue(Double.Parse(gridShowData_Driver1.Rows[j].Cells[k].Text.Trim()));
                            wsExcel.GetRow(vLinesNo).GetCell(k).CellStyle = csData;
                        }
                    }
                }
                vLinesNo++;
                wsExcel.CreateRow(vLinesNo);
                for (int i = 0; i < gridShowData_Driver1.Columns.Count; i++)
                {
                    vHeaderText = gridShowData_Driver1.Columns[i].HeaderText.Trim();
                    if (vHeaderText == "單位")
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(i);
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue("");
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    else if (vHeaderText == "列標籤")
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(i);
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue("總計");
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    else
                    {
                        vStrCode_1 = (i % 26) + 65;
                        vStrCode_2 = (i / 26);
                        vCellName = (vStrCode_2 != 0) ? ((char)(vStrCode_2 + 64)).ToString() + ((char)vStrCode_1).ToString() : ((char)vStrCode_1).ToString();

                        vCellName_S = vCellName + "2";
                        vCellName_E = vCellName + vLinesNo.ToString();
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("sum(" + vCellName_S + ":" + vCellName_E + ")");
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                }
            }
            vSheetName = "內勤統計";
            //新增一個工作表
            wsExcel = (XSSFSheet)wbExcel.CreateSheet(vSheetName);
            //重新取一次資料
            GetData_Emp(vMode);
            //寫入標題列
            vLinesNo = 0;
            wsExcel.CreateRow(vLinesNo);
            for (int i = 0; i < dtOverDutyList.Columns.Count; i++)
            {
                vHeaderText = dtOverDutyList.Columns[i].Caption.Trim();
                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
            }
            for (int i = 0; i < dtOverDutyList.Rows.Count; i++)
            {
                vLinesNo++;
                wsExcel.CreateRow(vLinesNo);
                for (int j = 0; j < dtOverDutyList.Columns.Count; j++)
                {
                    vHeaderText = dtOverDutyList.Columns[j].Caption.Trim();
                    if (vHeaderText == "單位")
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(j);
                        wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.String);
                        wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(dtOverDutyList.Rows[i][j].ToString().Trim());
                        wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = csData;
                    }
                    else
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(j);
                        wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.Numeric);
                        wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(Double.Parse(dtOverDutyList.Rows[i][j].ToString().Trim()));
                        wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = csData;
                    }
                }
            }
            vLinesNo++;
            wsExcel.CreateRow(vLinesNo);
            for (int i = 0; i < dtOverDutyList.Columns.Count; i++)
            {
                vHeaderText = dtOverDutyList.Columns[i].Caption.Trim();
                if (vHeaderText == "單位")
                {
                    wsExcel.GetRow(vLinesNo).CreateCell(i);
                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue("總計");
                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                }
                else
                {
                    vStrCode_1 = (i % 26) + 65;
                    vStrCode_2 = (i / 26);
                    vCellName = (vStrCode_2 != 0) ? ((char)(vStrCode_2 + 64)).ToString() + ((char)vStrCode_1).ToString() : ((char)vStrCode_1).ToString();

                    vCellName_S = vCellName + "2";
                    vCellName_E = vCellName + vLinesNo.ToString();
                    wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("sum(" + vCellName_S + ":" + vCellName_E + ")");
                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
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
                    string vCalYMStr = eCalYear.Text.Trim() + " 年 " + eCalMonth.Text.Trim() + " 月";
                    string vSplitModeStr = rbSplitMode.SelectedItem.Text.Trim();
                    string vRecordNote = "匯出檔案_工作會報用加班時數統計" + Environment.NewLine +
                                         "WorkReportOverTimes.aspx" + Environment.NewLine +
                                         "計算年月：" + vCalYMStr + Environment.NewLine +
                                         "區分性別：" + vSplitModeStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);

                    //先判斷是不是IE
                    HttpContext.Current.Response.ContentType = "application/octet-stream";
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
                Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}