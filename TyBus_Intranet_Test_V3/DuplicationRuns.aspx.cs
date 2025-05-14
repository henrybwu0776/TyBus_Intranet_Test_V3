using Amaterasu_Function;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class DuplicationRuns : System.Web.UI.Page
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
        //private string vTempStr = "";
        private DataTable dtShowData;

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

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo = eDepNo_Search.Text.Trim();
            string vTempStr = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
            string vDepName = PF.GetValue(vConnStr, vTempStr, "Name");
            if (vDepName == "")
            {
                vDepName = vDepNo.Trim();
                vTempStr = "select top 1 DepNo from Department where [Name] = '" + vDepName + "' order by DepNo";
                vDepNo = PF.GetValue(vConnStr, vTempStr, "DepNo");
            }
            eDepNo_Search.Text = vDepNo;
            eDepName_Search.Text = vDepName;
            if (eDepNo_Search.Text.Trim() != "")
            {
                using (SqlConnection connGetRunturn = new SqlConnection(vConnStr))
                {
                    vTempStr = "select ClassNo, ClassTxt from Dbdicb where FKEY = '勤務班次        turntime        RUNTURN' and ClassNo like '" + eDepNo_Search.Text.Trim() + "%' order by ClassNo";
                    SqlCommand cmdGetRunturn = new SqlCommand(vTempStr, connGetRunturn);
                    connGetRunturn.Open();
                    SqlDataReader drGetRunturn = cmdGetRunturn.ExecuteReader();
                    if (drGetRunturn.HasRows)
                    {
                        ddlRunturn_Search.Items.Clear();
                        ddlTargetRunturn_Search.Items.Clear();
                        while (drGetRunturn.Read())
                        {
                            ListItem iTemp = new ListItem(drGetRunturn["ClassTxt"].ToString().Trim(), drGetRunturn["ClassNo"].ToString().Trim());
                            ddlRunturn_Search.Items.Add(iTemp);
                            ddlTargetRunturn_Search.Items.Add(iTemp);
                        }
                    }
                }
            }
            else
            {
                ddlRunturn_Search.Items.Clear();
                ddlTargetRunturn_Search.Items.Clear();
            }
        }

        /// <summary>
        /// 顯示指定班次資訊
        /// </summary>
        /// <param name="fRunturnNo"></param>
        private void ShowData(string fRunturnNo)
        {
            string vSelectStr = "select d.[Name] 站別, r.Runturn 班次代碼, dd.ClassTxt 班次, ddd.ClassTxt 車種, rt.RunItemName 班次大類, rb.RunItemBName 班次小類, " + Environment.NewLine +
                                "       l.LineName 路線名稱, r.ToLine 去程路線, r.ToTime 去程時間, r.BackLine 回程路線, r.BackTime 回程時間, r.ToRemark 去程打卡, r.BackRemark 回程打卡 " + Environment.NewLine +
                                "  from Runs r left join Department d on d.DepNo = r.DepNo " + Environment.NewLine +
                                "              left join RunItem rt on rt.DepNo = r.DepNo and rt.RunItemNo = r.RunItem " + Environment.NewLine +
                                "              left join RunItemB rb on rb.RunItemBNo = r.RunItemB " + Environment.NewLine +
                                "              left join Lines l on l.LinesNo = r.LinesNo " + Environment.NewLine +
                                "              left join DBDICB dd on dd.FKey = '勤務班次        turntime        RUNTURN' and dd.ClassNo = r.Runturn " + Environment.NewLine +
                                "              left join DBDICB ddd on ddd.FKey = '班次表          Runs            CARTYPE' and ddd.ClassNo = r.CarType " + Environment.NewLine +
                                " where r.Runturn = @Runturn " + Environment.NewLine +
                                " order by RunitemB, ToTime ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connShowData = new SqlConnection(vConnStr))
            {
                if (dtShowData == null)
                {
                    dtShowData = new DataTable();
                }
                else
                {
                    dtShowData.Clear();
                }
                SqlDataAdapter daShowData = new SqlDataAdapter(vSelectStr, connShowData);
                connShowData.Open();
                daShowData.SelectCommand.Parameters.Clear();
                daShowData.SelectCommand.Parameters.Add(new SqlParameter("Runturn", fRunturnNo));
                daShowData.Fill(dtShowData);
                gvRunsList.DataSource = dtShowData;
                gvRunsList.DataBind();
            }
        }

        /// <summary>
        /// 搜尋班次資訊
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSearch_Click(object sender, EventArgs e)
        {
            string vRunturn_Temp = ddlRunturn_Search.SelectedValue.Trim();
            ShowData(vRunturn_Temp);
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCopy_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSourceRunturn = ddlRunturn_Search.SelectedValue.Trim();
            string vTargetRunturn = ddlTargetRunturn_Search.SelectedValue.Trim();
            string vTempStr = "";
            string vCopyDateStr = DateTime.Today.ToString("yyyyMMdd");
            //先把要刪除的班次複製到備份資料表
            vTempStr = "insert into RUNSBAK ( " + Environment.NewLine +
                       "       RunBakKey, BackupDate, RunItem, RunItemB, Cartype, ToTime, BackLine, LinesNo, ActualKM, ToRemark, " + Environment.NewLine +
                       "       BackTime, ToLine, BackRemark, Car_ID, Car_NO, Depno, RunTurn, EndTime, RunPK, IsExtra, IsFeed, " + Environment.NewLine +
                       "       StartDate, EndDate, remark, number, RunItemName, RunItemBName, ToShow, BackShow, PreName, CoercePrint, ReturnTime) " + Environment.NewLine +
                       "select @CopyDate + RunPK, GetDate(), RunItem, RunItemB, Cartype, ToTime, BackLine, LinesNo, ActualKM, ToRemark, " + Environment.NewLine +
                       "       BackTime, ToLine, BackRemark, Car_ID, Car_NO, Depno, RunTurn, EndTime, RunPK, IsExtra, IsFeed, " + Environment.NewLine +
                       "       StartDate, EndDate, remark, number, RunItemName, RunItemBName, ToShow, BackShow, PreName, CoercePrint, ReturnTime " + Environment.NewLine +
                       "  from Runs " + Environment.NewLine +
                       " where RunTurn = @TargetRunturn ";
            using (SqlConnection connBackup = new SqlConnection(vConnStr))
            {
                SqlCommand cmdBackup = new SqlCommand(vTempStr, connBackup);
                connBackup.Open();
                cmdBackup.Parameters.Clear();
                cmdBackup.Parameters.Add(new SqlParameter("CopyDate", vCopyDateStr));
                cmdBackup.Parameters.Add(new SqlParameter("TargetRunturn", vTargetRunturn));
                cmdBackup.ExecuteNonQuery();
            }
            //刪除指定的班次
            vTempStr = "delete Runs where Runturn = @TargetRunturn ";
            using (SqlConnection connDel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdDel = new SqlCommand(vTempStr, connDel);
                connDel.Open();
                cmdDel.Parameters.Clear();
                cmdDel.Parameters.Add(new SqlParameter("TargetRunturn", vTargetRunturn));
                cmdDel.ExecuteNonQuery();
            }
            vTempStr = "insert into Runs " + Environment.NewLine +
                       "       (RunItem, RunItemB, CarType, ToTime, BackLine, LinesNo, ActualKM, ToRemark, BackTime, ToLine, BackRemark, Car_ID, Car_No, DepNo, Runturn, EndTime, " + Environment.NewLine +
                       "        RunPK, IsExtra, IsFeed, StartDate, EndDate, Remark, Number, RunItemName, RunItemBName, ToShow, BackShow, PreName, CoercePrint, ReturnTime)" + Environment.NewLine +
                       "select RunItem, RunItemB, CarType, ToTime, BackLine, LinesNo, ActualKM, ToRemark, BackTime, ToLine, BackRemark, Car_ID, Car_No, DepNo, " + Environment.NewLine +
                       "       @TargetRunturn, EndTime, (@TargetRunturn + RunItemB + ToTime + convert(varchar(10), StartDate, 111)), IsExtra, IsFeed, " + Environment.NewLine +
                       "       StartDate, EndDate, Remark, Number, RunItemName, RunItemBName, ToShow, BackShow, PreName, CoercePrint, ReturnTime " + Environment.NewLine +
                       "  from Runs " + Environment.NewLine +
                       " where Runturn = @SourceRunturn ";
            using (SqlConnection connCopyData = new SqlConnection(vConnStr))
            {
                SqlCommand cmdCopyData = new SqlCommand(vTempStr, connCopyData);
                connCopyData.Open();
                cmdCopyData.Parameters.Clear();
                cmdCopyData.Parameters.Add(new SqlParameter("TargetRunturn", vTargetRunturn));
                cmdCopyData.Parameters.Add(new SqlParameter("SourceRunturn", vSourceRunturn));
                cmdCopyData.ExecuteNonQuery();
            }
            ShowData(vTargetRunturn);
        }
    }
}