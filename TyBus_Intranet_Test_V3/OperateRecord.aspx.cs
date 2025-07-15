using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class OperateRecord : System.Web.UI.Page
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
                vComputerName = Environment.MachineName; //2021.09.27 新增取得電腦名稱

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }

                    //事件日期
                    string vOperateDateURL = "InputDate.aspx?TextboxID=" + eOperateDateS_Search.ClientID;
                    string vOperateDateScript = "window.open('" + vOperateDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eOperateDateS_Search.Attributes["onClick"] = vOperateDateScript;

                    vOperateDateURL = "InputDate.aspx?TextboxID=" + eOperateDateE_Search.ClientID;
                    vOperateDateScript = "window.open('" + vOperateDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eOperateDateE_Search.Attributes["onClick"] = vOperateDateScript;

                    if (!IsPostBack)
                    {
                        plShowData_List.Visible = false;
                        plShowData_Detail.Visible = false;
                    }
                    else
                    {
                        OpenData();
                        plShowData_List.Visible = true;
                        plShowData_Detail.Visible = true;
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

        private string GetSelStr()
        {
            string vResultStr = "";
            string vWStr_LoginNo = (eLoginID_Search.Text.Trim() != "") ? "   and LoginID = '" + eLoginID_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_OperateDate = ((eOperateDateS_Search.Text.Trim() != "") && (eOperateDateE_Search.Text.Trim() != "")) ?
                                       "   and (OperateDate >= '" + eOperateDateS_Search.Text.Trim() + " 00:00:00' and OperateDate <= '" + eOperateDateE_Search.Text.Trim() + " 23:59:59') " + Environment.NewLine :
                                       ((eOperateDateS_Search.Text.Trim() != "") && (eOperateDateE_Search.Text.Trim() == "")) ?
                                       "   and OperateDate = '" + eOperateDateS_Search.Text.Trim() + "' " + Environment.NewLine :
                                       ((eOperateDateS_Search.Text.Trim() == "") && (eOperateDateE_Search.Text.Trim() != "")) ?
                                       "   and OperateDate = '" + eOperateDateE_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vOperateMode = (eOperateMode_Search.Text.Trim() != "") ? eOperateMode_Search.Text.Trim() : ddlOperateMode_Saerch.SelectedValue.Trim();
            string vWStr_Action = (vOperateMode == "00") ? "" :
                                  (vOperateMode == "01") ? "   and charindex('查詢資料', OperationNote, 0) > 0 " + Environment.NewLine :
                                  (vOperateMode == "02") ? "   and charindex('匯出檔案', OperationNote, 0) > 0 " + Environment.NewLine :
                                  (vOperateMode == "03") ? "   and (charindex('列印', OperationNote, 0) > 0 or charindex('預覽報表', OperationNote, 0) > 0) " + Environment.NewLine :
                                  (vOperateMode == "04") ? "   and charindex('查詢資料',OperationNote,0) = 0 and charindex('匯出檔案',OperationNote,0) = 0 and charindex('列印',OperationNote,0) = 0 and charindex('預覽報表', OperationNote, 0) = 0" + Environment.NewLine : "";
            string vWStr_FunctionName = ((ddlOperateFunctionMain_Search.SelectedValue != "") && (ddlOperateFunctionSub_Search.SelectedValue != "")) ?
                                        "   and charindex('" + ddlOperateFunctionSub_Search.SelectedValue.Trim() + "', OperationNote, 0) > 0 " + Environment.NewLine :
                                        ((ddlOperateFunctionMain_Search.SelectedValue != "") && (ddlOperateFunctionSub_Search.SelectedValue == "")) ?
                                        GetNoteStr() :
                                        (ddlOperateFunctionMain_Search.SelectedValue != "") ? "" : "";
            vResultStr = "SELECT IndexNo, LoginID, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = OperateRecord.LoginID)) AS EmpName, " + Environment.NewLine +
                         "       ComputerName, OperateDate, OperationNote " + Environment.NewLine +
                         "  FROM OperateRecord " + Environment.NewLine +
                         " WHERE (ISNULL(IndexNo, '') <> '') " + Environment.NewLine +
                         vWStr_Action +
                         vWStr_LoginNo +
                         vWStr_FunctionName +
                         vWStr_OperateDate +
                         " ORDER BY IndexNo DESC";
            return vResultStr;
        }

        private string GetNoteStr()
        {
            string vResultStr = "";
            string vTempStr = "";
            string vGroupID = ddlOperateFunctionMain_Search.SelectedValue.Trim();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                vTempStr = "select TargetPage from WebPermissionA where GroupID = '" + vGroupID + "' ";
                SqlCommand cmdTemp = new SqlCommand(vTempStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    while (drTemp.Read())
                    {
                        vResultStr = (vResultStr == "") ?
                                     "   and (charindex('" + drTemp[0].ToString().Trim() + "', OperationNote, 0) > 0 " + Environment.NewLine :
                                     vResultStr += "        or charindex('" + drTemp[0].ToString().Trim() + "', OperationNote, 0) > 0 " + Environment.NewLine;
                    }
                    vResultStr += ")" + Environment.NewLine;
                }
            }
            return vResultStr;
        }

        private void OpenData()
        {
            string vSelStr = GetSelStr();
            dsOperateRecord_List.SelectCommand = vSelStr;
            gridOperateRecord_List.DataSourceID = "";
            gridOperateRecord_List.DataSource = dsOperateRecord_List;
            gridOperateRecord_List.DataBind();
        }

        protected void eLoginID_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo = eLoginID_Search.Text.Trim();
            string vTempStr = "select [Name] from Employee where EmpNo = '" + vEmpNo + "' ";
            string vEmpName = PF.GetValue(vConnStr, vTempStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vTempStr = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName + "' order by EmpNo DESC";
                vEmpNo = PF.GetValue(vConnStr, vTempStr, "EmpNo");
            }
            eLoginID_Search.Text = vEmpNo.Trim();
            eLoginName_Search.Text = vEmpName.Trim();
        }

        protected void ddlOperateMode_Saerch_SelectedIndexChanged(object sender, EventArgs e)
        {
            eOperateMode_Search.Text = ddlOperateMode_Saerch.SelectedValue.Trim();
        }

        protected void ddlOperateFunctionMain_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            ddlOperateFunctionSub_Search.Items.Clear();
            string vGroupID_Search = ddlOperateFunctionMain_Search.SelectedValue.Trim();
            string vTempStr = "SELECT CAST(NULL AS varchar) AS GroupID, CAST(NULL AS varchar) AS TargetPage, CAST(NULL AS varchar) AS ControlCName " + Environment.NewLine +
                              " UNION ALL " + Environment.NewLine +
                              "SELECT GroupID, TargetPage, ControlCName " + Environment.NewLine +
                              "  FROM WebPermissionA " + Environment.NewLine +
                              " WHERE GroupID = '" + vGroupID_Search + "' " + Environment.NewLine +
                              " ORDER BY GroupID, TargetPage";
            dsOperateFunctionSub_Search.SelectCommand = vTempStr;
            ddlOperateFunctionSub_Search.DataSourceID = "";
            ddlOperateFunctionSub_Search.DataSource = dsOperateFunctionSub_Search;
            ddlOperateFunctionSub_Search.DataTextField = "ControlCName";
            ddlOperateFunctionSub_Search.DataValueField = "TargetPage";
            ddlOperateFunctionSub_Search.DataBind();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
            string vLoginIDStr = (eLoginID_Search.Text.Trim() != "") ? eLoginID_Search.Text.Trim() : "全部";
            string vOperateModeStr = (eOperateMode_Search.Text.Trim() != "") ? ddlOperateMode_Saerch.SelectedItem.Text.Trim() : "全部";
            string vOperateDateStr = ((eOperateDateS_Search.Text.Trim() != "") && (eOperateDateE_Search.Text.Trim() != "")) ? " 自 " + eOperateDateS_Search.Text.Trim() + " 起至 " + eOperateDateE_Search.Text.Trim() + " 止 " + Environment.NewLine :
                                     ((eOperateDateS_Search.Text.Trim() != "") && (eOperateDateE_Search.Text.Trim() == "")) ? " 於 " + eOperateDateS_Search.Text.Trim() + Environment.NewLine :
                                     ((eOperateDateS_Search.Text.Trim() == "") && (eOperateDateE_Search.Text.Trim() != "")) ? "於 " + eOperateDateE_Search.Text.Trim() + Environment.NewLine : "不指定日期區間";
            string vOperateFunctionStr = ((ddlOperateFunctionMain_Search.SelectedValue != "") && (ddlOperateFunctionSub_Search.SelectedValue != "")) ?
                                         ddlOperateFunctionMain_Search.SelectedItem.Text + "--" + ddlOperateFunctionSub_Search.SelectedItem.Text :
                                         ((ddlOperateFunctionMain_Search.SelectedValue != "") && (ddlOperateFunctionSub_Search.SelectedValue == "")) ?
                                         ddlOperateFunctionMain_Search.SelectedItem.Text : "全部";
            string vRecordNote = "查詢資料_系統操作記錄" + Environment.NewLine +
                                 "OperateRecord.aspx" + Environment.NewLine +
                                 "登入帳號：" + vLoginIDStr + Environment.NewLine +
                                 "操作行為：" + vOperateModeStr + Environment.NewLine +
                                 "操作日期時間：" + vOperateDateStr + Environment.NewLine +
                                 "操作功能：" + vOperateFunctionStr;
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
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
            string vSelStr = GetSelStr();
            string vFileName = "線上系統操作記錄清冊";
            DateTime vBuDate;

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
                        vHeaderText = (drExcel.GetName(i).Trim().ToUpper() == "INDEXNO") ? "序號" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "LOGINID") ? "登入帳號" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "EMPNAME") ? "姓名" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "COMPUTERNAME") ? "電腦名稱 (IP)" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "OPERATEDATE") ? "操作日期時間" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "OPERATIONNOTE") ? "操作行為" : drExcel.GetName(i).Trim();
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
                            if ((drExcel.GetName(i).Trim().ToUpper() == "OPERATEDATE") && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.ToString("yyyy/MM/dd HH:mm:dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
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
                            string vLoginIDStr = (eLoginID_Search.Text.Trim() != "") ? eLoginID_Search.Text.Trim() : "全部";
                            string vOperateModeStr = (eOperateMode_Search.Text.Trim() != "") ? ddlOperateMode_Saerch.SelectedItem.Text.Trim() : "全部";
                            string vOperateDateStr = ((eOperateDateS_Search.Text.Trim() != "") && (eOperateDateE_Search.Text.Trim() != "")) ? " 自 " + eOperateDateS_Search.Text.Trim() + " 起至 " + eOperateDateE_Search.Text.Trim() + " 止 " + Environment.NewLine :
                                                     ((eOperateDateS_Search.Text.Trim() != "") && (eOperateDateE_Search.Text.Trim() == "")) ? " 於 " + eOperateDateS_Search.Text.Trim() + Environment.NewLine :
                                                     ((eOperateDateS_Search.Text.Trim() == "") && (eOperateDateE_Search.Text.Trim() != "")) ? "於 " + eOperateDateE_Search.Text.Trim() + Environment.NewLine : "不指定日期區間";
                            string vOperateFunctionStr = ((ddlOperateFunctionMain_Search.SelectedValue != "") && (ddlOperateFunctionSub_Search.SelectedValue != "")) ?
                                                         ddlOperateFunctionMain_Search.SelectedItem.Text + "--" + ddlOperateFunctionSub_Search.SelectedItem.Text :
                                                         ((ddlOperateFunctionMain_Search.SelectedValue != "") && (ddlOperateFunctionSub_Search.SelectedValue == "")) ?
                                                         ddlOperateFunctionMain_Search.SelectedItem.Text : "全部";
                            string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                 "OperateRecord.aspx" + Environment.NewLine +
                                                 "登入帳號：" + vLoginIDStr + Environment.NewLine +
                                                 "操作行為：" + vOperateModeStr + Environment.NewLine +
                                                 "操作日期時間：" + vOperateDateStr + Environment.NewLine +
                                                 "操作功能：" + vOperateFunctionStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
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
                        Response.Write("alert('" + eMessage.Message + "')");
                        Response.Write("</" + "Script>");
                    }
                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void gridOperateRecord_List_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridOperateRecord_List.PageIndex = e.NewPageIndex;
            gridOperateRecord_List.DataBind();
        }
    }
}