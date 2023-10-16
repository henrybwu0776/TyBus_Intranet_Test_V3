using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class PermissionSetting : System.Web.UI.Page
    {
        PublicFunction PF = new PublicFunction(); //加入公用程式碼參考
        private string vLoginID = "";
        private string vDepNo = "";
        private string vConnStr = "";
        DateTime vToday = DateTime.Today;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"] != null)
            {
                vLoginID = (Session["LoginID"] != null) ? Session["LoginID"].ToString().Trim() : "";
                vDepNo = (Session["LoginDepNo"] != null) ? Session["LoginDepNo"].ToString().Trim() : "";

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        if (ddlGroupID_Search.Items.Count > 0)
                        {
                            ddlGroupID_Search.SelectedIndex = 0;
                            eGroupID_Search.Text = ddlGroupID_Search.Items[0].Value.Trim();
                        }
                    }
                    PermissionA_Databind();
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
            string vWStr_DepNo = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? "   and ControlName in (select DISTINCT ControlName from WebPermissionB where DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "') " + Environment.NewLine :
                                 ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? "   and ControlName in (select DISTINCT ControlName from WebPermissionB where DepNo = '" + eDepNo_Start_Search.Text.Trim() + "') " + Environment.NewLine :
                                 ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? "   and ControlName in (select DISTINCT ControlName from WebPermissionB where DepNo = '" + eDepNo_End_Search.Text.Trim() + "') " + Environment.NewLine : "";
            string vWStr_EmpNo = ((eEmpNo_Start_Search.Text.Trim() != "") && (eEmpNo_End_Search.Text.Trim() != "")) ? "   and ControlName in (select DISTINCT ControlName from WebPermissionB where EmpNo between '" + eEmpNo_Start_Search.Text.Trim() + "' and '" + eEmpNo_End_Search.Text.Trim() + "') " + Environment.NewLine :
                                 ((eEmpNo_Start_Search.Text.Trim() != "") && (eEmpNo_End_Search.Text.Trim() == "")) ? "   and ControlName in (select DISTINCT ControlName from WebPermissionB where EmpNo = '" + eEmpNo_Start_Search.Text.Trim() + "') " + Environment.NewLine :
                                 ((eEmpNo_Start_Search.Text.Trim() == "") && (eEmpNo_End_Search.Text.Trim() != "")) ? "   and ControlName in (select DISTINCT ControlName from WebPermissionB where EmpNo = '" + eEmpNo_End_Search.Text.Trim() + "') " + Environment.NewLine : "";
            string vWStr_GroupID = (eGroupID_Search.Text.Trim() != "") ? "   and GroupID = '" + eGroupID_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ControlCName = (eControlCName_Search.Text.Trim() != "") ? "   and ControlCName like '%" + eControlCName_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vSQLStr = "select ControlName, OrderIndex, ControlCName, TargetPage, Remark, " + Environment.NewLine +
                             "       GroupID, (select ClassTxt from DBDICB where ClassNo = a.GroupID and FKey = '網頁功能權限群組WebPermission   GroupID') GroupName " + Environment.NewLine +
                             "  from WebPermissionA a where 1 = 1 " + Environment.NewLine +
                             vWStr_DepNo + vWStr_EmpNo + vWStr_GroupID + vWStr_ControlCName +
                             " order by OrderIndex ";
            return vSQLStr;
        }

        private void PermissionA_Databind()
        {
            sdsWebPermissionA_List.SelectCommand = "";
            sdsWebPermissionA_List.SelectCommand = GetSelectStr();
            gridWebPermissionA_List.DataSourceID = "sdsWebPermissionA_List";
            sdsWebPermissionA_List.DataBind();
        }

        protected void ddlDepNo_Start_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDepNo_Start_Search.Text = ddlDepNo_Start.SelectedValue.Trim();
        }

        protected void ddlDepNo_End_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDepNo_End_Search.Text = ddlDepNo_End.SelectedValue.Trim();
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            HSSFWorkbook wbExcel = new HSSFWorkbook();
            //Excel 工作表
            HSSFSheet wsExcel;
            //設定標題欄位的格式
            HSSFCellStyle csTitle = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中

            //設定字體格式
            HSSFFont fontTitle = (HSSFFont)wbExcel.CreateFont();
            //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
            fontTitle.IsBold = true;
            fontTitle.FontHeightInPoints = 12; //字體大小
            csTitle.SetFont(fontTitle);

            //設定資料內容欄位的格式
            HSSFCellStyle csData = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            HSSFFont fontData = (HSSFFont)wbExcel.CreateFont();
            fontData.FontHeightInPoints = 12;
            csData.SetFont(fontData);

            HSSFCellStyle csData_Red = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            HSSFFont fontData_Red = (HSSFFont)wbExcel.CreateFont();
            fontData_Red.FontHeightInPoints = 12; //字體大小
            fontData_Red.Color = HSSFColor.Red.Index; //字體顏色
            csData_Red.SetFont(fontData_Red);

            HSSFCellStyle csData_Int = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            HSSFDataFormat format = (HSSFDataFormat)wbExcel.CreateDataFormat();
            csData_Int.DataFormat = format.GetFormat("###,##0");

            string vHeaderText = "";
            int vLinesNo = 0;
            string vFileName = "桃園客運內部線上系統功能清單";
            string vSelStr = GetSelectStr();

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i) == "ControlName") ? "功能名稱" :
                                      (drExcel.GetName(i) == "OrderIndex") ? "排列順序" :
                                      (drExcel.GetName(i) == "ControlCName") ? "功能中文名" :
                                      (drExcel.GetName(i) == "TargetPage") ? "目標網頁" :
                                      (drExcel.GetName(i) == "Remark") ? "備註說明" :
                                      (drExcel.GetName(i) == "GroupID") ? "所屬群組代號" :
                                      (drExcel.GetName(i) == "GroupName") ? "所屬群組" : drExcel.GetName(i);
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    vLinesNo++;
                    while (drExcel.Read())
                    {
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                        }
                        vLinesNo++;
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
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                            {
                                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xls", System.Text.Encoding.UTF8));
                            }
                            else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                            {
                                // 設定強制下載標頭
                                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName + ".xls"));
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
                        Response.Write("alert('" + eMessage.Message.Trim() + "')");
                        Response.Write("</" + "Script>");
                        //throw;
                    }
                }
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            PermissionA_Databind();
        }

        protected void BatchNewButton_EmptyB_Click(object sender, EventArgs e)
        {
            Label lbTempControlName_BatchNew_EmptyB = (Label)fvPermissionA_Data.FindControl("eControlName_ListA");
            string vTempControlName = lbTempControlName_BatchNew_EmptyB.Text.Trim();
            string vTempItemStr = "";
            string vTempControlNameItems = "";
            string vTempDepNo = "";
            string vTempAllowPermission = "1";
            string vTempRemark = "批次授權";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            string vSQLStr = "select DISTINCT DepNo from Employee where LeaveDay is NULL and isnull(DepNo, '00') <> '00'";
            SqlConnection connTemp = new SqlConnection(vConnStr);
            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
            connTemp.Open();
            SqlDataReader drTemp = cmdTemp.ExecuteReader();
            while (drTemp.Read())
            {
                vSQLStr = "select Max(Items) MaxItem from WebPermissionB where ControlName = '" + vTempControlName + "' ";
                vTempItemStr = PF.GetValue(vConnStr, vSQLStr, "MaxItem");
                vTempItemStr = (vTempItemStr != "") ? (int.Parse(vTempItemStr) + 1).ToString("D4") : "0001";
                vTempControlNameItems = vTempControlName + vTempItemStr;
                vTempDepNo = drTemp["DepNo"].ToString();
                vSQLStr = "insert into WebPermissionB" + Environment.NewLine +
                          "       (ControlName, Items, ControlNameItems, DepNo, EmpNo, AllowPermission, RemarkB)" + Environment.NewLine +
                          " values('" + vTempControlName + "', '" + vTempItemStr + "', '" + vTempControlNameItems + "', '" + vTempDepNo + "'," + Environment.NewLine +
                          "        NULL, '" + vTempAllowPermission + "', '" + vTempRemark + "')";
                PF.ExecSQL(vConnStr, vSQLStr);
            }
        }

        protected void sdsWebPermissionA_Data_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            //A檔資料完成刪除後
            if (e.Exception == null)
            {
                this.gridWebPermissionA_List.DataBind();
            }
        }

        protected void sdsWebPermissionA_Data_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            //A檔資料完成新增後
            if (e.Exception == null)
            {
                this.gridWebPermissionA_List.DataBind();
            }
        }

        protected void sdsWebPermissionA_Data_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            //A檔資料完成修改後
            if (e.Exception == null)
            {
                this.gridWebPermissionA_List.DataBind();
            }
        }

        protected void sdsWebPermissionA_Data_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            //A檔資料開始新增前
            Label eOrderIndex_InsertA = (Label)fvPermissionA_Data.FindControl("eOrderIndex_InsertA");
            string vTempStr = eOrderIndex_InsertA.Text.Trim();
            if (vTempStr == "")
            {
                string vSQLStr = "select Max(OrderIndex) MaxIndex from WebPermissionA";
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                string vTempOrderIndex = PF.GetValue(vConnStr, vSQLStr, "MaxIndex");
                e.Command.Parameters["@OrderIndex"].Value = (vTempOrderIndex != "") ? int.Parse(vTempOrderIndex) + 1 : 1;
            }
        }

        protected void sdsWebPermissionB_Data_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            //B檔資料刪除後
            if (e.Exception == null)
            {
                this.gridWebPermissionB_List.DataBind();
            }
        }

        protected void sdsWebPermissionB_Data_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            //B檔資料新增後
            if (e.Exception == null)
            {
                this.gridWebPermissionB_List.DataBind();
            }
        }

        protected void sdsWebPermissionB_Data_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            //B檔資料開始新增前
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label lbTempControlName_InsertB = (Label)fvPermissionA_Data.FindControl("eControlName_ListA");
            string vTempControlName = lbTempControlName_InsertB.Text.Trim();
            string vSQLStr = "select Max(Items) MaxItem from WebPermissionB where ControlName = '" + vTempControlName + "' ";
            string vTempItemStr = PF.GetValue(vConnStr, vSQLStr, "MaxItem");
            vTempItemStr = (vTempItemStr != "") ? (int.Parse(vTempItemStr) + 1).ToString("D4") : "0001";
            string vTempControlNameItems = vTempControlName + vTempItemStr;
            e.Command.Parameters["@ControlName"].Value = vTempControlName;
            e.Command.Parameters["@Items"].Value = vTempItemStr;
            e.Command.Parameters["@ControlNameItems"].Value = vTempControlNameItems;
        }

        protected void sdsWebPermissionB_Data_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            //B檔資料修改後
            if (e.Exception == null)
            {
                this.gridWebPermissionB_List.DataBind();
            }
        }

        protected void ddlGroupID_InsertA_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlGroupID_InsertA = (DropDownList)fvPermissionA_Data.FindControl("ddlGroupID_InsertA");
            TextBox eGroupID_InsertA = (TextBox)fvPermissionA_Data.FindControl("eGroupID_InsertA");
            eGroupID_InsertA.Text = ddlGroupID_InsertA.SelectedValue.Trim();
        }

        protected void ddlGroupID_EditA_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlGroupID_EditA = (DropDownList)fvPermissionA_Data.FindControl("ddlGroupID_EditA");
            TextBox eGroupID_EditA = (TextBox)fvPermissionA_Data.FindControl("eGroupID_EditA");
            eGroupID_EditA.Text = ddlGroupID_EditA.SelectedValue.Trim();
        }

        protected void fvPermissionA_Data_DataBound(object sender, EventArgs e)
        {
            switch (fvPermissionA_Data.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    DropDownList ddlGroupID_EditA = (DropDownList)fvPermissionA_Data.FindControl("ddlGroupID_EditA");
                    TextBox eGroupID_EditA = (TextBox)fvPermissionA_Data.FindControl("eGroupID_EditA");
                    ddlGroupID_EditA.SelectedIndex = ddlGroupID_EditA.Items.IndexOf(ddlGroupID_EditA.Items.FindByValue(eGroupID_EditA.Text.Trim()));
                    break;
                case FormViewMode.Insert:
                    break;
            }
        }

        protected void ddlGroupID_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eGroupID_Search.Text = (ddlGroupID_Search.SelectedValue != "") ? ddlGroupID_Search.SelectedValue.Trim() : ddlGroupID_Search.Items[0].Value.Trim();
        }

        protected void eDepNo_InsertB_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_InsertB = (TextBox)fvPermissionB_Data.FindControl("eDepNo_InsertB");
            if (eDepNo_InsertB != null)
            {
                Label eDepName_InsertB = (Label)fvPermissionB_Data.FindControl("eDepName_InsertB");
                string vDepNo_Temp = eDepNo_InsertB.Text.Trim();
                string vDepName_Temp = "";
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
                vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName_Temp == "")
                {
                    vDepName_Temp = vDepNo_Temp.Trim();
                    vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                    vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo_InsertB.Text = vDepNo_Temp.Trim();
                eDepName_InsertB.Text = vDepName_Temp.Trim();
            }
        }

        protected void eEmpNo_InsertB_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eEmpNo_InsertB = (TextBox)fvPermissionB_Data.FindControl("eEmpNo_InsertB");
            if (eEmpNo_InsertB != null)
            {
                Label eEmpName_InsertB = (Label)fvPermissionB_Data.FindControl("eEmpName_InsertB");
                string vEmpNo_Temp = eEmpNo_InsertB.Text.Trim();
                string vEmpName_Temp = "";
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp.Trim() + "' ";
                vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vEmpName_Temp == "")
                {
                    vEmpName_Temp = vEmpNo_Temp.Trim();
                    vSQLStr_Temp = "select EmpNo from Employee where [Name] = '" + vEmpName_Temp.Trim() + "' ";
                    vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eEmpNo_InsertB.Text = vEmpNo_Temp.Trim();
                eEmpName_InsertB.Text = vEmpName_Temp.Trim();
            }
        }

        protected void eDepNo_EditB_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_EditB = (TextBox)fvPermissionB_Data.FindControl("eDepNo_EditB");
            if (eDepNo_EditB != null)
            {
                Label eDepName_EditB = (Label)fvPermissionB_Data.FindControl("eDepName_EditB");
                string vDepNo_Temp = eDepNo_EditB.Text.Trim();
                string vDepName_Temp = "";
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
                vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName_Temp == "")
                {
                    vDepName_Temp = vDepNo_Temp.Trim();
                    vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                    vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo_EditB.Text = vDepNo_Temp.Trim();
                eDepName_EditB.Text = vDepName_Temp.Trim();
            }
        }

        protected void eEmpNo_EditB_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eEmpNo_EditB = (TextBox)fvPermissionB_Data.FindControl("eEmpNo_EditB");
            if (eEmpNo_EditB != null)
            {
                Label eEmpName_EditB = (Label)fvPermissionB_Data.FindControl("eEmpName_EditB");
                string vEmpNo_Temp = eEmpNo_EditB.Text.Trim();
                string vEmpName_Temp = "";
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp.Trim() + "' ";
                vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vEmpName_Temp == "")
                {
                    vEmpName_Temp = vEmpNo_Temp.Trim();
                    vSQLStr_Temp = "select EmpNo from Employee where [Name] = '" + vEmpName_Temp.Trim() + "' ";
                    vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eEmpNo_EditB.Text = vEmpNo_Temp.Trim();
                eEmpName_EditB.Text = vEmpName_Temp.Trim();
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbClearRight_Click(object sender, EventArgs e)
        {
            string vSQLStr_Clear = "";
            string vMessageStr = "本功能只能一次清除一個員工的權限，" + Environment.NewLine +
                                 "多名員工需清除時請分次進行。" + Environment.NewLine +
                                 "不支援單位權限清除！";
            Response.Write("<Script language='Javascript'>");
            Response.Write("alert('" + vMessageStr + "')");
            Response.Write("</" + "Script>");
            eDepNo_Start_Search.Text = "";
            eDepNo_End_Search.Text = "";
            eEmpNo_End_Search.Text = "";
            string vEmpNo_Clear = eEmpNo_Start_Search.Text.Trim();
            if (vEmpNo_Clear != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                vSQLStr_Clear = "update WebPermissionB set AllowPermission = 0, " + Environment.NewLine +
                                "       RemarkB = RemarkB + " + Environment.NewLine + "convert(varchar(10),getdate(),111) + '批次修正' " + Environment.NewLine +
                                " where EmpNo = '" + vEmpNo_Clear + "' ";
                PF.ExecSQL(vConnStr, vSQLStr_Clear);
            }
        }

        /// <summary>
        /// 從 EXCEL 批次更新權限
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbImportRight_Click(object sender, EventArgs e)
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName);
                string vTempStr = "";
                int vSheetCount = 0;

                string vEmpNo_Excel = "";
                string vControlName_Excel = "";
                string vAllowPermission_Excel = "";
                string vDepNo_Excel = "";
                string vItems_Excel = "";
                string vMaxItems = "";
                string vRemarkB_Excel = "";

                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }

                switch (vExtName)
                {
                    case ".xls":
                        //Excel 97-2003
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        vSheetCount = wbExcel_H.NumberOfSheets;
                        for (int vSC = 0; vSC < vSheetCount; vSC++)
                        {
                            HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(vSC);
                            for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                            {
                                HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                                vControlName_Excel = vRowTemp_H.GetCell(0).ToString().Trim();
                                if ((vControlName_Excel != "") && (vControlName_Excel.ToUpper() != "CONTROLNAME"))
                                {
                                    vAllowPermission_Excel = vRowTemp_H.GetCell(1).ToString().Trim();
                                    vEmpNo_Excel = vRowTemp_H.GetCell(2).ToString().Trim();
                                    vDepNo_Excel = vRowTemp_H.GetCell(3).ToString().Trim();
                                    vTempStr = "select MAX(Items) as MaxItems from WebPermissionB where ControlName = '" + vControlName_Excel + "' ";
                                    vMaxItems = PF.GetValue(vConnStr, vTempStr, "MaxItems");
                                    vItems_Excel = (vMaxItems.Trim() != "") ? (Int32.Parse(vMaxItems) + 1).ToString("D4") : "0001";
                                    vRemarkB_Excel = DateTime.Today.ToShortDateString() + "_EXCEL檔整批匯入";
                                    using (SqlDataSource sdsExcel = new SqlDataSource())
                                    {
                                        sdsExcel.ConnectionString = vConnStr;
                                        sdsExcel.InsertCommand = "insert into WebPermissionB (ControlName, Items, ControlNameItems, DepNo, EmpNo, AllowPermission, RemarkB)" + Environment.NewLine +
                                                                 " values (@ControlName, @Items, @ControlNameItems, @DepNo, @EmpNo, @AllowPermission, @RemarkB)";
                                        sdsExcel.InsertParameters.Clear();
                                        sdsExcel.InsertParameters.Add(new Parameter("ControlName", DbType.String, vControlName_Excel));
                                        sdsExcel.InsertParameters.Add(new Parameter("Items", DbType.String, vItems_Excel));
                                        sdsExcel.InsertParameters.Add(new Parameter("ControlNameItems", DbType.String, vControlName_Excel.Trim() + vItems_Excel.Trim()));
                                        sdsExcel.InsertParameters.Add(new Parameter("DepNo", DbType.String, (vEmpNo_Excel.Trim() == "") ? vDepNo_Excel : String.Empty));
                                        sdsExcel.InsertParameters.Add(new Parameter("EmpNo", DbType.String, (vEmpNo_Excel.Trim() != "") ? vEmpNo_Excel : String.Empty));
                                        sdsExcel.InsertParameters.Add(new Parameter("AllowPermission", DbType.Boolean, (vAllowPermission_Excel.Trim() == "1") ? "true" : "false"));
                                        sdsExcel.InsertParameters.Add(new Parameter("RemarkB", DbType.String, (vRemarkB_Excel.Trim() != "") ? vRemarkB_Excel : String.Empty));
                                        sdsExcel.Insert();
                                    }
                                }
                            }
                        }
                        break;

                    case ".xlsx":
                        //Excel 2010--
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        vSheetCount = wbExcel_X.NumberOfSheets;
                        for (int vSC = 0; vSC < vSheetCount; vSC++)
                        {
                            XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(vSC);
                            for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
                            {
                                XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                                vControlName_Excel = vRowTemp_X.GetCell(0).ToString().Trim();
                                if ((vControlName_Excel != "") && (vControlName_Excel.ToUpper() != "CONTROLNAME"))
                                {
                                    vAllowPermission_Excel = vRowTemp_X.GetCell(1).ToString().Trim();
                                    vEmpNo_Excel = vRowTemp_X.GetCell(2).ToString().Trim();
                                    vDepNo_Excel = vRowTemp_X.GetCell(3).ToString().Trim();
                                    vTempStr = "select MAX(Items) as MaxItems from WebPermissionB where ControlName = '" + vControlName_Excel + "' ";
                                    vMaxItems = PF.GetValue(vConnStr, vTempStr, "MaxItems");
                                    vItems_Excel = (vMaxItems.Trim() != "") ? (Int32.Parse(vMaxItems) + 1).ToString("D4") : "0001";
                                    vRemarkB_Excel = DateTime.Today.ToShortDateString() + "_EXCEL檔整批匯入";
                                    using (SqlDataSource sdsExcel = new SqlDataSource())
                                    {
                                        sdsExcel.ConnectionString = vConnStr;
                                        sdsExcel.InsertCommand = "insert into WebPermissionB (ControlName, Items, ControlNameItems, DepNo, EmpNo, AllowPermission, RemarkB)" + Environment.NewLine +
                                                                 " values (@ControlName, @Items, @ControlNameItems, @DepNo, @EmpNo, @AllowPermission, @RemarkB)";
                                        sdsExcel.InsertParameters.Clear();
                                        sdsExcel.InsertParameters.Add(new Parameter("ControlName", DbType.String, vControlName_Excel));
                                        sdsExcel.InsertParameters.Add(new Parameter("Items", DbType.String, vItems_Excel));
                                        sdsExcel.InsertParameters.Add(new Parameter("ControlNameItems", DbType.String, vControlName_Excel.Trim() + vItems_Excel.Trim()));
                                        sdsExcel.InsertParameters.Add(new Parameter("DepNo", DbType.String, (vEmpNo_Excel.Trim() == "") ? vDepNo_Excel : String.Empty));
                                        sdsExcel.InsertParameters.Add(new Parameter("EmpNo", DbType.String, (vEmpNo_Excel.Trim() != "") ? vEmpNo_Excel : String.Empty));
                                        sdsExcel.InsertParameters.Add(new Parameter("AllowPermission", DbType.Boolean, (vAllowPermission_Excel.Trim() == "1") ? "true" : "false"));
                                        sdsExcel.InsertParameters.Add(new Parameter("RemarkB", DbType.String, (vRemarkB_Excel.Trim() != "") ? vRemarkB_Excel : String.Empty));
                                        sdsExcel.Insert();
                                    }
                                }
                            }
                        }
                        break;
                }
            }
        }
    }
}