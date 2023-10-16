using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class LinesNoChart : Page
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
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        
                    }
                    else
                    {
                        OpenDataList();
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
            string vWStr_ERPLinesNo = (eERPLinesNo_Search.Text.Trim() != "") ? "   and ERPLinesNo like '" + eERPLinesNo_Search.Text.Trim() + "%' " : "";
            string vWStr_GovLinesNo = ((eGOVLinesNo_Search.Text.Trim() != "") && (eGOVLinesExtNo_Search.Text.Trim() != "")) ? "   and GOVLinesNo like '" + eGOVLinesNo_Search.Text.Trim() + "%' and GOVLinesExtNo = '" + eGOVLinesExtNo_Search.Text.Trim() + "' " :
                                      ((eGOVLinesNo_Search.Text.Trim() != "") && (eGOVLinesExtNo_Search.Text.Trim() == "")) ? "   and GOVLinesNo like '" + eGOVLinesNo_Search.Text.Trim() + "%' " :
                                      ((eGOVLinesNo_Search.Text.Trim() == "") && (eGOVLinesExtNo_Search.Text.Trim() != "")) ? "   and GOVLinesExtNo = '" + eGOVLinesExtNo_Search.Text.Trim() + "' " : "";
            string vWStr_TicketLinesNo = (eTicketLinesNo_Search.Text.Trim() != "") ? "   and TicketLinesNo like '" + eTicketLinesNo_Search.Text.Trim() + "%' " : "";
            string vResultStr = "SELECT IndexNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, " + Environment.NewLine +
                                "       ERPLinesNo, GOVLinesNo, GOVLinesExtNo, TicketLinesNo, Remark " + Environment.NewLine +
                                "  FROM LinesNoChart AS a " + Environment.NewLine +
                                " WHERE isnull(IndexNo, '') <> '' " + Environment.NewLine +
                                vWStr_ERPLinesNo + vWStr_GovLinesNo + vWStr_TicketLinesNo +
                                " order by ERPLinesNo";
            return vResultStr;
        }

        private void OpenDataList()
        {
            string vSelStr = GetSelectStr();
            sdsLinesNoChart_List.SelectCommand = vSelStr;
            gridLinesNoChartList.DataBind();
            //fvLinesNoChartDetail.DataBind();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenDataList();
        }

        /// <summary>
        /// 轉入路線對照資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            string vExtName = "";
            string vTempStr = "";
            int vRC = 0;
            string vMaxIndexNoStr = "";
            string vErrorStr = "";

            string vIndexNo = "";
            string vMaxIndexNo = "";
            string vIndexNoHead = "";
            string vTicketLineNo = "";
            string vERPLineNo = "";
            string vLineGovNo = "";
            string vDepNo = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            if (fuExcel.FileName != "")
            {
                vExtName = Path.GetExtension(fuExcel.FileName);
                switch (vExtName)
                {
                    default:
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('您指定的檔案不是 EXCEL 檔，請重新確認檔案格式。')");
                        Response.Write("</" + "Script>");
                        break;

                    case ".xls": //EXCEL 2003 之前版本
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet vSheet_H = (HSSFSheet)wbExcel_H.GetSheet("分支路線展開");
                        if (vSheet_H.LastRowNum > 0)
                        {
                            for (int vRCount = vSheet_H.FirstRowNum + 1; vRCount <= vSheet_H.LastRowNum; vRCount++)
                            {
                                try
                                {
                                    HSSFRow vRowTemp_H = (HSSFRow)vSheet_H.GetRow(vRCount);
                                    vTicketLineNo = vRowTemp_H.GetCell(2).ToString().Trim();
                                    vERPLineNo = (vRowTemp_H.GetCell(4).ToString().Trim() != "-") ? Int32.Parse(vRowTemp_H.GetCell(4).ToString().Trim()).ToString("D5") : vERPLineNo;
                                    vLineGovNo = vRowTemp_H.GetCell(0).ToString().Trim();

                                    vTempStr = "select count(LinesNo) RCount from Lines where LinesNo = '" + vERPLineNo + "' ";
                                    vRC = Int32.Parse(PF.GetValue(vConnStr, vTempStr, "RCount"));
                                    if (vRC == 0)
                                    {
                                        vErrorStr += (Environment.NewLine + "系統內找不到ERP路線代號 [" + vERPLineNo + "]，請重新確認！");
                                    }
                                    else if (vRC == 1)
                                    {
                                        vTempStr = "select top 1 DepNo from Department where [Name] = '" + vRowTemp_H.GetCell(10).ToString().Trim() + "' ";
                                        vDepNo = PF.GetValue(vConnStr, vTempStr, "DepNo");
                                        vTempStr = "select count(TicketLinesNo) RCount from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                        vRC = Int32.Parse(PF.GetValue(vConnStr, vTempStr, "RCount"));
                                        if (vRC == 0)
                                        {
                                            vIndexNoHead = DateTime.Today.ToString("yyyyMM") + "ERPLines";
                                            vTempStr = "select Max(IndexNo) MaxIndex from LinesNoChart where IndexNo like '" + vIndexNoHead + "%' ";
                                            vMaxIndexNoStr = PF.GetValue(vConnStr, vTempStr, "MaxIndex").Replace(vIndexNoHead, "");
                                            vMaxIndexNo = (vMaxIndexNoStr != "") ? (Int32.Parse(vMaxIndexNoStr) + 1).ToString("D4") : "0001";
                                            vIndexNo = vIndexNoHead.Trim() + vMaxIndexNo;
                                            vTempStr = "INSERT INTO LinesNoChart " + Environment.NewLine +
                                                       "       (IndexNo, DepNo, ERPLinesNo, GOVLinesNo, TicketLinesNo, BuMan, BuDate) " + Environment.NewLine +
                                                       "VALUES ('" + vIndexNo + "', '" + vDepNo + "', '" + vERPLineNo + "', '" + vLineGovNo + "', '" + vTicketLineNo + "', '" + vLoginID + "', GetDate())";
                                            PF.ExecSQL(vConnStr, vTempStr);
                                        }
                                        else if (vRC == 1)
                                        {
                                            vTempStr = "select IndexNo from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                            vIndexNo = PF.GetValue(vConnStr, vTempStr, "IndexNo");
                                            vTempStr = "UPDATE LinesNoChart " + Environment.NewLine +
                                                       "   SET DepNo = '" + vDepNo + "', ERPLinesNo = '" + vERPLineNo + "', GOVLinesNo = '" + vLineGovNo + "', " + Environment.NewLine +
                                                       "       TicketLinesNo = '" + vTicketLineNo + "', ModifyMan = '" + vLoginID + "', ModifyDate = GetDate() " + Environment.NewLine +
                                                       " WHERE (IndexNo = '" + vIndexNo + "' )";
                                            PF.ExecSQL(vConnStr, vTempStr);
                                        }
                                        else
                                        {

                                        }
                                        vTempStr = "update Lines " + Environment.NewLine +
                                                   "   set TicketLineNo = '" + Int32.Parse(vTicketLineNo).ToString("D6") + "', LinesGOVNo = '" + vLineGovNo + "', ApprovedDepNo = '" + vDepNo + "' " + Environment.NewLine +
                                                   " where LinesNo = '" + vERPLineNo + "' ";
                                        PF.ExecSQL(vConnStr, vTempStr);
                                    }
                                    else
                                    {

                                    }
                                }
                                catch (Exception eMessage)
                                {
                                    Response.Write("<Script language='Javascript'>");
                                    Response.Write("alert('" + eMessage.Message + "')");
                                    Response.Write("</" + "Script>");
                                }
                            }
                        }
                        break;

                    case ".xlsx": //EXCEL 2010 之後版本
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet vSheet_X = (XSSFSheet)wbExcel_X.GetSheet("分支路線展開");
                        if (vSheet_X.LastRowNum > 0)
                        {
                            for (int vRCount = vSheet_X.FirstRowNum + 1; vRCount <= vSheet_X.LastRowNum; vRCount++)
                            {
                                try
                                {
                                    XSSFRow vRowTemp_X = (XSSFRow)vSheet_X.GetRow(vRCount);
                                    vTicketLineNo = vRowTemp_X.GetCell(2).ToString().Trim();
                                    vERPLineNo = (vRowTemp_X.GetCell(4).ToString().Trim() != "-") ? Int32.Parse(vRowTemp_X.GetCell(4).ToString().Trim()).ToString("D5") : vERPLineNo;
                                    vLineGovNo = vRowTemp_X.GetCell(0).ToString().Trim();

                                    vTempStr = "select count(LinesNo) RCount from Lines where LinesNo = '" + vERPLineNo + "' ";
                                    vRC = Int32.Parse(PF.GetValue(vConnStr, vTempStr, "RCount"));
                                    if (vRC == 0)
                                    {
                                        vErrorStr += (Environment.NewLine + "系統內找不到ERP路線代號 [" + vERPLineNo + "]，請重新確認！");
                                    }
                                    else if (vRC == 1)
                                    {
                                        vTempStr = "select top 1 DepNo from Department where [Name] = '" + vRowTemp_X.GetCell(10).ToString().Trim() + "' ";
                                        vDepNo = PF.GetValue(vConnStr, vTempStr, "DepNo");
                                        vTempStr = "select count(TicketLinesNo) RCount from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                        vRC = Int32.Parse(PF.GetValue(vConnStr, vTempStr, "RCount"));
                                        if (vRC == 0)
                                        {
                                            vIndexNoHead = DateTime.Today.ToString("yyyyMM") + "ERPLines";
                                            vTempStr = "select Max(IndexNo) MaxIndex from LinesNoChart where IndexNo like '" + vIndexNoHead + "%' ";
                                            vMaxIndexNoStr = PF.GetValue(vConnStr, vTempStr, "MaxIndex").Replace(vIndexNoHead, "");
                                            vMaxIndexNo = (vMaxIndexNoStr != "") ? (Int32.Parse(vMaxIndexNoStr) + 1).ToString("D4") : "0001";
                                            vIndexNo = vIndexNoHead.Trim() + vMaxIndexNo;
                                            vTempStr = "INSERT INTO LinesNoChart " + Environment.NewLine +
                                                       "       (IndexNo, DepNo, ERPLinesNo, GOVLinesNo, TicketLinesNo, BuMan, BuDate) " + Environment.NewLine +
                                                       "VALUES ('" + vIndexNo + "', '" + vDepNo + "', '" + vERPLineNo + "', '" + vLineGovNo + "', '" + vTicketLineNo + "', '" + vLoginID + "', GetDate())";
                                            PF.ExecSQL(vConnStr, vTempStr);
                                        }
                                        else if (vRC == 1)
                                        {
                                            vTempStr = "select IndexNo from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                            vIndexNo = PF.GetValue(vConnStr, vTempStr, "IndexNo");
                                            vTempStr = "UPDATE LinesNoChart " + Environment.NewLine +
                                                       "   SET DepNo = '" + vDepNo + "', ERPLinesNo = '" + vERPLineNo + "', GOVLinesNo = '" + vLineGovNo + "', " + Environment.NewLine +
                                                       "       TicketLinesNo = '" + vTicketLineNo + "', ModifyMan = '" + vLoginID + "', ModifyDate = GetDate() " + Environment.NewLine +
                                                       " WHERE (IndexNo = '" + vIndexNo + "' )";
                                            PF.ExecSQL(vConnStr, vTempStr);
                                        }
                                        else
                                        {

                                        }
                                        vTempStr = "update Lines " + Environment.NewLine +
                                                   "   set TicketLineNo = '" + Int32.Parse(vTicketLineNo).ToString("D6") + "', LinesGOVNo = '" + vLineGovNo + "', ApprovedDepNo = '" + vDepNo + "' " + Environment.NewLine +
                                                   " where LinesNo = '" + vERPLineNo + "' ";
                                        PF.ExecSQL(vConnStr, vTempStr);
                                    }
                                    else
                                    {

                                    }
                                }
                                catch (Exception eMessage)
                                {
                                    Response.Write("<Script language='Javascript'>");
                                    Response.Write("alert('" + eMessage.Message + Environment.NewLine + vTempStr + "')");
                                    Response.Write("</" + "Script>");
                                }
                            }
                        }
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

        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            Label eIndexNo_Edit = (Label)fvLinesNoChartDetail.FindControl("eIndexNo_Edit");
            if (eIndexNo_Edit != null)
            {
                if (eIndexNo_Edit.Text.Trim() != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    TextBox eDepNo_Edit = (TextBox)fvLinesNoChartDetail.FindControl("eDepNo_Edit");
                    TextBox eERPLinesNo_Edit = (TextBox)fvLinesNoChartDetail.FindControl("eERPLinesNo_Edit");
                    TextBox eGOVLinesNo_Edit = (TextBox)fvLinesNoChartDetail.FindControl("eGOVLinesNo_Edit");
                    TextBox eGOVLinesExtNo_Edit = (TextBox)fvLinesNoChartDetail.FindControl("eGOVLinesExtNo_Edit");
                    TextBox eTicketLinesNo_Edit = (TextBox)fvLinesNoChartDetail.FindControl("eTicketLinesNo_Edit");
                    TextBox eRemark_Edit = (TextBox)fvLinesNoChartDetail.FindControl("eRemark_Edit");

                    try
                    {
                        //先複製一份到異動資料表
                        string vTempStr = "select MAX(Items) MaxItem from LinesNoChart_History where IndexNo = '" + eIndexNo_Edit.Text.Trim() + "' ";
                        string vMaxItem = PF.GetValue(vConnStr, vTempStr, "MaxItem");
                        int vNewItem = (vMaxItem != "") ? Int32.Parse(vMaxItem) + 1 : 1;
                        string vNewItemStr = vNewItem.ToString("D4");
                        string vIndexNoItem = eIndexNo_Edit.Text.Trim() + vNewItemStr.Trim();
                        vTempStr = "insert into LinesNoChart_History " + Environment.NewLine +
                                   "       (IndexNoItem, Items, ModifyMode, IndexNo, DepNo, " + Environment.NewLine +
                                   "        ERPLinesNo, GOVLinesNo, GOVLinesExtNo, TicketLinesNo, " + Environment.NewLine +
                                   "        Remark, BuMan, BuDate, ModifyMan, ModifyDate) " + Environment.NewLine +
                                   " select '" + vIndexNoItem + "', '" + vNewItemStr + "', '修改', IndexNo, DepNo, ERPLinesNo, GOVLinesNo, GOVLinesExtNo, TicketLinesNo, " + Environment.NewLine +
                                   "        Remark, BuMan, BuDate, '" + vLoginID + "', '" + DateTime.Today.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                                   "   from LinesNoChart " + Environment.NewLine +
                                   "  where IndexNo = '" + eIndexNo_Edit.Text.Trim() + "' ";
                        PF.ExecSQL(vConnStr, vTempStr);

                        //把異動後的資料寫入資料表
                        sdsLinesNoChart_Detail.UpdateParameters.Clear();
                        sdsLinesNoChart_Detail.UpdateCommand = "UPDATE LinesNoChart " + Environment.NewLine +
                                                               "   SET DepNo = @DepNo, ERPLinesNo = @ERPLinesNo, GOVLinesNo = @GOVLinesNo, " + Environment.NewLine +
                                                               "       GOVLinesExtNo = @GOVLinesExtNo, TicketLinesNo = @TicketLinesNo, Remark = @Remark, " + Environment.NewLine +
                                                               "       ModifyMan = @ModifyMan, ModifyDate = @ModifyDate " + Environment.NewLine +
                                                               " WHERE (IndexNo = @IndexNo)";
                        sdsLinesNoChart_Detail.UpdateParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo_Edit.Text.Trim() != "") ? eDepNo_Edit.Text.Trim() : String.Empty));
                        sdsLinesNoChart_Detail.UpdateParameters.Add(new Parameter("ERPLinesNo", DbType.String, (eERPLinesNo_Edit.Text.Trim() != "") ? eERPLinesNo_Edit.Text.Trim() : String.Empty));
                        sdsLinesNoChart_Detail.UpdateParameters.Add(new Parameter("GOVLinesNo", DbType.String, (eGOVLinesNo_Edit.Text.Trim() != "") ? eGOVLinesNo_Edit.Text.Trim() : String.Empty));
                        sdsLinesNoChart_Detail.UpdateParameters.Add(new Parameter("GOVLinesExtNo", DbType.String, (eGOVLinesExtNo_Edit.Text.Trim() != "") ? eGOVLinesExtNo_Edit.Text.Trim() : String.Empty));
                        sdsLinesNoChart_Detail.UpdateParameters.Add(new Parameter("TicketLinesNo", DbType.String, (eTicketLinesNo_Edit.Text.Trim() != "") ? eTicketLinesNo_Edit.Text.Trim() : String.Empty));
                        sdsLinesNoChart_Detail.UpdateParameters.Add(new Parameter("Remark", DbType.String, (eRemark_Edit.Text.Trim() != "") ? eRemark_Edit.Text.Trim() : String.Empty));
                        sdsLinesNoChart_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                        sdsLinesNoChart_Detail.UpdateParameters.Add(new Parameter("ModifyDate", DbType.Date, DateTime.Today.ToString("yyyy/MM/dd")));
                        sdsLinesNoChart_Detail.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, eIndexNo_Edit.Text.Trim()));
                        sdsLinesNoChart_Detail.Update();
                        fvLinesNoChartDetail.ChangeMode(FormViewMode.ReadOnly);
                        gridLinesNoChartList.DataBind();
                        fvLinesNoChartDetail.DataBind();
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
                    Response.Write("alert('請先選擇要修改的路線代號！')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Edit = (TextBox)fvLinesNoChartDetail.FindControl("eDepNo_Edit");
            if (eDepNo_Edit != null)
            {
                try
                {
                    Label eDepName_Edit = (Label)fvLinesNoChartDetail.FindControl("eDepName_Edit");
                    string vDepNo_Edit = eDepNo_Edit.Text.Trim();
                    string vTempStr = "select [Name] from Department where DepNo = '" + vDepNo_Edit + "' ";
                    string vDepName_Edit = PF.GetValue(vConnStr, vTempStr, "Name");
                    if (vDepName_Edit == "")
                    {
                        vDepName_Edit = vDepNo_Edit.Trim();
                        vTempStr = "select top 1 DepNo from Department where [Name] = '" + vDepName_Edit + "' ";
                        vDepNo_Edit = PF.GetValue(vConnStr, vTempStr, "DepNo");
                    }
                    eDepName_Edit.Text = vDepName_Edit.Trim();
                    eDepNo_Edit.Text = vDepNo_Edit.Trim();
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
            Label eIndexNo_INS = (Label)fvLinesNoChartDetail.FindControl("eIndexNo_INS");
            if (eIndexNo_INS != null)
            {
                if (vConnStr=="")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vIndexNo_INS = "";
                string vIndexNoHead = DateTime.Today.ToString("yyyyMM") + "ERPLines";
                string vTempStr = "select Max(IndexNo) MaxIndex from LinesNoChart where IndexNo like '" + vIndexNoHead + "%' ";
                string vMaxIndexNo = PF.GetValue(vConnStr, vTempStr, "MaxIndex");
                int vNewIndex = (vMaxIndexNo.Trim() != "") ? Int32.Parse(vMaxIndexNo.Replace(vIndexNoHead, "")) + 1 : 1;
                vIndexNo_INS = vIndexNoHead + vNewIndex.ToString("D4");

                TextBox eDepNo_INS = (TextBox)fvLinesNoChartDetail.FindControl("eDepNo_INS");
                TextBox eERPLinesNo_INS = (TextBox)fvLinesNoChartDetail.FindControl("eERPLinesNo_INS");
                TextBox eGOVLinesNo_INS = (TextBox)fvLinesNoChartDetail.FindControl("eGOVLinesNo_INS");
                TextBox eGOVLinesExtNo_INS = (TextBox)fvLinesNoChartDetail.FindControl("eGOVLinesExtNo_INS");
                TextBox eTicketLinesNo_INS = (TextBox)fvLinesNoChartDetail.FindControl("eTicketLinesNo_INS");
                TextBox eRemark_INS = (TextBox)fvLinesNoChartDetail.FindControl("eRemark_INS");
                try
                {
                    sdsLinesNoChart_Detail.InsertParameters.Clear();
                    sdsLinesNoChart_Detail.InsertCommand = "INSERT INTO LinesNoChart " + Environment.NewLine +
                                                           "       (IndexNo, DepNo, ERPLinesNo, GOVLinesNo, GOVLinesExtNo, TicketLinesNo, Remark, BuMan, BuDate) " + Environment.NewLine +
                                                           "VALUES (@IndexNo, @DepNo, @ERPLinesNo, @GOVLinesNo, @GOVLinesExtNo, @TicketLinesNo, @Remark, @BuMan, @BuDate)";
                    sdsLinesNoChart_Detail.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo_INS));
                    sdsLinesNoChart_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo_INS.Text.Trim() != "") ? eDepNo_INS.Text.Trim() : String.Empty));
                    sdsLinesNoChart_Detail.InsertParameters.Add(new Parameter("ERPLinesNo", DbType.String, (eERPLinesNo_INS.Text.Trim() != "") ? eERPLinesNo_INS.Text.Trim() : String.Empty));
                    sdsLinesNoChart_Detail.InsertParameters.Add(new Parameter("GOVLinesNo", DbType.String, (eGOVLinesNo_INS.Text.Trim() != "") ? eGOVLinesNo_INS.Text.Trim() : String.Empty));
                    sdsLinesNoChart_Detail.InsertParameters.Add(new Parameter("GOVLinesExtNo", DbType.String, (eGOVLinesExtNo_INS.Text.Trim() != "") ? eGOVLinesExtNo_INS.Text.Trim() : String.Empty));
                    sdsLinesNoChart_Detail.InsertParameters.Add(new Parameter("TicketLinesNo", DbType.String, (eTicketLinesNo_INS.Text.Trim() != "") ? eTicketLinesNo_INS.Text.Trim() : String.Empty));
                    sdsLinesNoChart_Detail.InsertParameters.Add(new Parameter("Remark", DbType.String, (eRemark_INS.Text.Trim() != "") ? eRemark_INS.Text.Trim() : String.Empty));
                    sdsLinesNoChart_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    sdsLinesNoChart_Detail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, DateTime.Today.ToString("yyyy/MM/dd")));
                    sdsLinesNoChart_Detail.Insert();
                    fvLinesNoChartDetail.ChangeMode(FormViewMode.ReadOnly);
                    gridLinesNoChartList.DataBind();
                    fvLinesNoChartDetail.DataBind();
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

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_INS = (TextBox)fvLinesNoChartDetail.FindControl("eDepNo_INS");
            if (eDepNo_INS != null)
            {
                try
                {
                    Label eDepName_INS = (Label)fvLinesNoChartDetail.FindControl("eDepName_INS");
                    string vDepNo_INS = eDepNo_INS.Text.Trim();
                    string vTempStr = "select [Name] from Department where DepNo = '" + vDepNo_INS + "' ";
                    string vDepName_INS = PF.GetValue(vConnStr, vTempStr, "Name");
                    if (vDepName_INS == "")
                    {
                        vDepName_INS = vDepNo_INS.Trim();
                        vTempStr = "select top 1 DepNo from Department where [Name] = '" + vDepName_INS + "' ";
                        vDepNo_INS = PF.GetValue(vConnStr, vTempStr, "DepNo");
                    }
                    eDepName_INS.Text = vDepName_INS.Trim();
                    eDepNo_INS.Text = vDepNo_INS.Trim();
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
            Label eIndexNo_List = (Label)fvLinesNoChartDetail.FindControl("eIndexNo_List");
            if ((eIndexNo_List != null) && (eIndexNo_List.Text.Trim() != ""))
            {
                try
                {
                    string vIndexNo = eIndexNo_List.Text.Trim();
                    //先複製一份到異動資料表
                    string vTempStr = "select MAX(Items) MaxItem from LinesNoChart_History where IndexNo = '" + vIndexNo + "' ";
                    string vMaxItem = PF.GetValue(vConnStr, vTempStr, "MaxItem");
                    int vNewItem = (vMaxItem != "") ? Int32.Parse(vMaxItem) + 1 : 1;
                    string vNewItemStr = vNewItem.ToString("D4");
                    string vIndexNoItem = vIndexNo + vNewItemStr.Trim();
                    vTempStr = "insert into LinesNoChart_History " + Environment.NewLine +
                               "       (IndexNoItem, Items, ModifyMode, IndexNo, DepNo, " + Environment.NewLine +
                               "        ERPLinesNo, GOVLinesNo, GOVLinesExtNo, TicketLinesNo, " + Environment.NewLine +
                               "        Remark, BuMan, BuDate, ModifyMan, ModifyDate) " + Environment.NewLine +
                               " select '" + vIndexNoItem + "', '" + vNewItemStr + "', '刪除', IndexNo, DepNo, ERPLinesNo, GOVLinesNo, GOVLinesExtNo, TicketLinesNo, " + Environment.NewLine +
                               "        Remark, BuMan, BuDate, '" + vLoginID + "', '" + DateTime.Today.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                               "   from LinesNoChart " + Environment.NewLine +
                               "  where IndexNo = '" + vIndexNo + "' ";
                    PF.ExecSQL(vConnStr, vTempStr);

                    //把資料刪除
                    sdsLinesNoChart_Detail.DeleteParameters.Clear();
                    sdsLinesNoChart_Detail.DeleteParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                    sdsLinesNoChart_Detail.Delete();
                    gridLinesNoChartList.DataBind();
                    fvLinesNoChartDetail.DataBind();
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

        protected void fvLinesNoChartDetail_DataBound(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvLinesNoChartDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Label eIndexNo_List = (Label)fvLinesNoChartDetail.FindControl("eIndexNo_List");
                    if (eIndexNo_List != null)
                    {
                        Button bbDelete_List = (Button)fvLinesNoChartDetail.FindControl("bbDelete_List");
                        bbDelete_List.Enabled = (eIndexNo_List.Text.Trim() != "");
                    }
                    break;
                case FormViewMode.Edit:
                    TextBox eDepNo_Edit = (TextBox)fvLinesNoChartDetail.FindControl("eDepNo_Edit");
                    if (eDepNo_Edit != null)
                    {
                        Label eDepName_Edit = (Label)fvLinesNoChartDetail.FindControl("eDepName_Edit");
                    }
                    break;
                case FormViewMode.Insert:
                    TextBox eDepNo_INS = (TextBox)fvLinesNoChartDetail.FindControl("eDepNo_INS");
                    if (eDepNo_INS != null)
                    {
                        Label eDepName_INS = (Label)fvLinesNoChartDetail.FindControl("eDepName_INS");
                    }
                    break;
                default:
                    break;
            }
        }
    }
}