using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class RetireFreeTicketList : System.Web.UI.Page
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
                        eTicketYear_Search.Text = "";
                        eRetireYear_Search.Text = "";
                        plShowData_Main.Visible = true;
                        plShowData_Detail.Visible = false;
                        plPrint.Visible = false;
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

        protected void eEmpNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo_Temp = eEmpNo_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp + "' ";
            string vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vEmpName_Temp.Trim() == "")
            {
                vEmpName_Temp = vEmpNo_Temp.Trim();
                vSQLStr_Temp = "select EmpNo from Employee where [Name] = '" + vEmpName_Temp + "' ";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eEmpNo_Search.Text = vEmpNo_Temp.Trim();
            eEmpName_Search.Text = vEmpName_Temp.Trim();
        }

        /// <summary>
        /// 取回查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            string vWStr_EmpNo = (eEmpNo_Search.Text.Trim() != "") ? "   and EmpNo = '" + eEmpNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_RetireYear = (eRetireYear_Search.Text.Trim() != "") ? "   and (year(a.RetireDate) - 1911) = " + eRetireYear_Search.Text.Trim() + Environment.NewLine : "";
            string vWStr_TicketYear = (eTicketYear_Search.Text.Trim() != "") ? "   and EmpNo in (select EmpNo from RetireFreeTicketB where TicketYear = '" + eTicketYear_Search.Text.Trim() + "')" + Environment.NewLine : "";
            string vResultStr = "select EmpNo, [Name], AssumeDate, RetireDate, " + Environment.NewLine +
                                "       TicketType, (select ClassTxt from DBDICB where FKey = '退休人員乘車證  RetireFreeTicketTicketType' and ClassNo = a.TicketType) TicketType_C, " + Environment.NewLine +
                                "       Remark, BuMan, (select [Name] from Employee where EmpNo = a.BuMan) BuManName, BuDate, " + Environment.NewLine +
                                "       ModifyMan, (select [Name] from Employee where EmpNo = a.ModifyMan) ModifyManName, ModifyDate " + Environment.NewLine +
                                "  from RetireFreeTicketA a " + Environment.NewLine +
                                " where isnull(a.EmpNo, '') <> '' " + Environment.NewLine +
                                vWStr_EmpNo + vWStr_RetireYear + vWStr_TicketYear +
                                " order by EmpNo, RetireDate DESC ";
            return vResultStr;
        }

        private void OpenData()
        {
            string vSelStr = GetSelectStr();
            sdsShowData_MainList.SelectCommand = vSelStr;
            gridShowData_Main.DataBind();
        }

        /// <summary>
        /// 查詢資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSerach_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        /// <summary>
        /// 取回匯出 EXCEL 用的查詢語法
        /// </summary>
        /// <returns></returns>
        private string GetSelStr_Excel()
        {
            string vWStr_EmpNo = (eEmpNo_Search.Text.Trim() != "") ? "   and EmpNo = '" + eEmpNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_RetireYear = (eRetireYear_Search.Text.Trim() != "") ? "   and (year(a.RetireDate) - 1911) = " + eRetireYear_Search.Text.Trim() + Environment.NewLine : "";
            string vWStr_TicketYear = (eTicketYear_Search.Text.Trim() != "") ? "   and b.TicketYear = '" + eTicketYear_Search.Text.Trim() + "')" + Environment.NewLine : "";
            string vResultStr = "select a.EmpNo, a.[Name], a.AssumeDate, a.RetireDate, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where ClassNo = a.TicketType and FKey = '退休人員乘車證  RetireFreeTicketTicketType') TicketType_C, " + Environment.NewLine +
                                "       Remark, b.Items, b.TicketYear, b.AssignDate, b.RemarkB " + Environment.NewLine +
                                "  from RetireFreeTicketA a left join RetireFreeTicketB b on b.EmpNo = a.EmpNo " + Environment.NewLine +
                                " where isnull(a.EmpNo, '') <> '' " + Environment.NewLine +
                                vWStr_EmpNo + vWStr_RetireYear + vWStr_TicketYear +
                                " order by a.EmpNo, b.Items";
            return vResultStr;
        }

        /// <summary>
        /// 匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
            string vFileName = "退休人員乘車證發放清冊";
            DateTime vBuDate;

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                string vSelStr = GetSelStr_Excel();
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
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "EMPNO") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "NAME") ? "員工姓名" :
                                      (drExcel.GetName(i).ToUpper() == "ASSUMEDATE") ? "到職日期" :
                                      (drExcel.GetName(i).ToUpper() == "RETIREDATE") ? "退休日期" :
                                      (drExcel.GetName(i).ToUpper() == "TICKETTYPE_C") ? "優待類別" :
                                      (drExcel.GetName(i).ToUpper() == "REMARK") ? "備註" :
                                      (drExcel.GetName(i).ToUpper() == "ITEMS") ? "項次" :
                                      (drExcel.GetName(i).ToUpper() == "TICKETYEAR") ? "適用年度" :
                                      (drExcel.GetName(i).ToUpper() == "ASSIGNDATE") ? "申辦日期" :
                                      (drExcel.GetName(i).ToUpper() == "REMARKB") ? "備註二" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i).ToUpper() == "ASSUMEDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "RETIREDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "ASSIGNDATE")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
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
                            string vEmpNoStr = (eEmpNo_Search.Text.Trim() != "") ? eEmpNo_Search.Text.Trim() : "全部";
                            string vRetireYearStr = (eRetireYear_Search.Text.Trim() != "") ? eRetireYear_Search.Text.Trim() : "全部";
                            string vTicketYearStr = (eTicketYear_Search.Text.Trim() != "") ? eTicketYear_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_退休人員乘車證維護作業" + Environment.NewLine +
                                                 "RetireFreeTicketList.aspx" + Environment.NewLine +
                                                 "員工工號：" + vEmpNoStr + Environment.NewLine +
                                                 "退休年度：" + vRetireYearStr + Environment.NewLine +
                                                 "乘車證年度：" + vTicketYearStr;
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
                        Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                        Response.Write("</" + "Script>");
                        //throw;
                    }
                }
            }
        }

        /// <summary>
        /// 結束回到主頁面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 主檔表格完成資料繫結
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void fvRetireFreeTicketA_DataBound(object sender, EventArgs e)
        {
            string vSQLStr_Temp = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvRetireFreeTicketA.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Label eEmpNo_List = (Label)fvRetireFreeTicketA.FindControl("eEmpNo_List");
                    Label eTicketType_List = (Label)fvRetireFreeTicketA.FindControl("eTicketType_List");
                    if (eEmpNo_List != null)
                    {
                        plShowData_Detail.Visible = ((eEmpNo_List.Text.Trim() != "") && (eTicketType_List.Text.Trim() != "0"));
                    }
                    else
                    {
                        plShowData_Detail.Visible = false;
                    }
                    break;
                case FormViewMode.Edit:
                    plShowData_Detail.Visible = false;
                    DropDownList ddlTicketType_Edit = (DropDownList)fvRetireFreeTicketA.FindControl("ddlTicketType_Edit");
                    Label eTicketType_Edit = (Label)fvRetireFreeTicketA.FindControl("eTicketType_Edit");
                    if (eTicketType_Edit != null)
                    {
                        ddlTicketType_Edit.Items.Clear();
                        vSQLStr_Temp = "SELECT ClassNo, CLASSTXT " + Environment.NewLine +
                                       "  FROM DBDICB " + Environment.NewLine +
                                       " WHERE (FKEY = '退休人員乘車證  RetireFreeTicketTicketType')";
                        using (SqlConnection connTemp_Edit = new SqlConnection(vConnStr))
                        {
                            SqlDataAdapter daTemp_Edit = new SqlDataAdapter(vSQLStr_Temp, connTemp_Edit);
                            connTemp_Edit.Open();
                            DataTable dtTemp_Edit = new DataTable();
                            daTemp_Edit.Fill(dtTemp_Edit);
                            ddlTicketType_Edit.DataSource = dtTemp_Edit;
                            ddlTicketType_Edit.DataValueField = "ClassNo";
                            ddlTicketType_Edit.DataTextField = "ClassTxt";
                            ddlTicketType_Edit.DataBind();
                            ddlTicketType_Edit.SelectedIndex = Int32.Parse(eTicketType_Edit.Text.Trim());
                        }
                    }
                    break;
                case FormViewMode.Insert:
                    plShowData_Detail.Visible = false;
                    DropDownList ddlTicketType_INS = (DropDownList)fvRetireFreeTicketA.FindControl("ddlTicketType_INS");
                    Label eTicketType_INS = (Label)fvRetireFreeTicketA.FindControl("eTicketType_INS");
                    if (eTicketType_INS != null)
                    {
                        ddlTicketType_INS.Items.Clear();
                        vSQLStr_Temp = "SELECT ClassNo, CLASSTXT " + Environment.NewLine +
                                       "  FROM DBDICB " + Environment.NewLine +
                                       " WHERE (FKEY = '退休人員乘車證  RetireFreeTicketTicketType')";
                        using (SqlConnection connTemp_INS = new SqlConnection(vConnStr))
                        {
                            SqlDataAdapter daTemp_INS = new SqlDataAdapter(vSQLStr_Temp, connTemp_INS);
                            connTemp_INS.Open();
                            DataTable dtTemp_INS = new DataTable();
                            daTemp_INS.Fill(dtTemp_INS);
                            ddlTicketType_INS.DataSource = dtTemp_INS;
                            ddlTicketType_INS.DataValueField = "ClassNo";
                            ddlTicketType_INS.DataTextField = "ClassTxt";
                            ddlTicketType_INS.DataBind();
                            ddlTicketType_INS.SelectedIndex = 0;
                            eTicketType_INS.Text = "0";
                        }
                    }
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// 主檔修改確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            string vSQLStr_Temp = "";
            string vMaxIndex = "";
            int vNewIndex = 0;
            string vItems = "";
            string vModifyMode = "EDIT";
            string vEmpNoItems = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eEmpNo_Edit = (Label)fvRetireFreeTicketA.FindControl("eEmpNo_Edit");
            try
            {
                if ((eEmpNo_Edit != null) && (eEmpNo_Edit.Text.Trim() != ""))
                {
                    //備份主檔
                    vSQLStr_Temp = "select max(Items) maxIndex from RetireFreeTicketA_History where EmpNo = '" + eEmpNo_Edit.Text.Trim() + "' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "maxIndex");
                    vNewIndex = (vMaxIndex.Trim() != "") ? Int32.Parse(vMaxIndex) + 1 : 1;
                    vItems = vNewIndex.ToString("D4");
                    vEmpNoItems = eEmpNo_Edit.Text.Trim() + vItems.Trim();
                    vSQLStr_Temp = "insert into RetireFreeTicketA_History " + Environment.NewLine +
                                   "       (EmpNoItem, ModifyMode, EmpNo, Items, Name, AssumeDate, RetireDate, TicketType, Remark, " + Environment.NewLine +
                                   "        BuMan, BuDate, ModifyMan, ModifyDate, CreateMan, CreateDate) " + Environment.NewLine +
                                   "select '" + vEmpNoItems + "', '" + vModifyMode + "', EmpNo, '" + vItems + "', Name, AssumeDate, RetireDate, TicketType, Remark, " + Environment.NewLine +
                                   "       BuMan, BuDate, ModifyMan, ModifyDate, '" + vLoginID + "', GetDate() " + Environment.NewLine +
                                   "  from RetireFreeTicketA " + Environment.NewLine +
                                   " where EmpNo = '" + eEmpNo_Edit.Text.Trim() + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);

                    //儲存修改
                    Label eTicketType_Edit = (Label)fvRetireFreeTicketA.FindControl("eTicketType_Edit");
                    TextBox eRemark_Edit = (TextBox)fvRetireFreeTicketA.FindControl("eRemark_Edit");
                    sdsShowData_Main.UpdateCommand = "UPDATE RetireFreeTicketA " + Environment.NewLine +
                                                     "   SET TicketType = @TicketType, Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate " + Environment.NewLine +
                                                     " WHERE (EmpNo = @EmpNo)";

                    sdsShowData_Main.UpdateParameters.Clear();
                    sdsShowData_Main.UpdateParameters.Add(new Parameter("TicketType", DbType.String, (eTicketType_Edit.Text.Trim() != "") ? eTicketType_Edit.Text.Trim() : String.Empty));
                    sdsShowData_Main.UpdateParameters.Add(new Parameter("Remark", DbType.String, (eRemark_Edit.Text.Trim() != "") ? eRemark_Edit.Text.Trim() : String.Empty));
                    sdsShowData_Main.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsShowData_Main.UpdateParameters.Add(new Parameter("ModifyDate", DbType.Date, DateTime.Today.ToShortDateString()));
                    sdsShowData_Main.UpdateParameters.Add(new Parameter("EmpNo", DbType.String, eEmpNo_Edit.Text.Trim()));

                    sdsShowData_Main.Update();
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('資料有誤！員工編號欄位異常！')");
                    Response.Write("</" + "Script>");
                }
                fvRetireFreeTicketA.ChangeMode(FormViewMode.ReadOnly);
                OpenData();
                fvRetireFreeTicketA.DataBind();
            }
            catch (Exception eMessage)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
        }

        /// <summary>
        /// 優待類別修改變更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlTicketType_Edit_TextChanged(object sender, EventArgs e)
        {
            DropDownList ddlTicketType_Temp = (DropDownList)fvRetireFreeTicketA.FindControl("ddlTicketType_Edit");
            Label eTicketType_Temp = (Label)fvRetireFreeTicketA.FindControl("eTicketType_Edit");
            if (ddlTicketType_Temp != null)
            {
                eTicketType_Temp.Text = ddlTicketType_Temp.SelectedValue.Trim();
                plShowData_Detail.Visible = (eTicketType_Temp.Text.Trim() != "0");
            }
        }

        /// <summary>
        /// 主檔新增確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eEmpNo_INS = (TextBox)fvRetireFreeTicketA.FindControl("eEmpNo_INS");
            if ((eEmpNo_INS != null) && (eEmpNo_INS.Text.Trim() != ""))
            {
                Label eEmpName_INS = (Label)fvRetireFreeTicketA.FindControl("eEmpName_INS");
                Label eAssumeDate_INS = (Label)fvRetireFreeTicketA.FindControl("eAssumeDate_INS");
                Label eRetireDate_INS = (Label)fvRetireFreeTicketA.FindControl("eRetireDate_INS");
                Label eTicketType_INS = (Label)fvRetireFreeTicketA.FindControl("eTicketType_INS");
                TextBox eRemark_INS = (TextBox)fvRetireFreeTicketA.FindControl("eRemark_INS");

                sdsShowData_Main.InsertCommand = "INSERT INTO RetireFreeTicketA " + Environment.NewLine +
                                                 "       (EmpNo, Name, AssumeDate, RetireDate, TicketType, Remark, BuMan, BuDate) " + Environment.NewLine +
                                                 "VALUES (@EmpNo, @Name, @AssumeDate, @RetireDate, @TicketType, @Remark, @BuMan, @BuDate)";
                sdsShowData_Main.InsertParameters.Clear();
                sdsShowData_Main.InsertParameters.Add(new Parameter("EmpNo", DbType.String, eEmpNo_INS.Text.Trim()));
                sdsShowData_Main.InsertParameters.Add(new Parameter("Name", DbType.String, (eEmpName_INS.Text.Trim() != "") ? eEmpName_INS.Text.Trim() : String.Empty));
                sdsShowData_Main.InsertParameters.Add(new Parameter("AssumeDate", DbType.Date, (eAssumeDate_INS.Text.Trim() != "") ? eAssumeDate_INS.Text.Trim() : String.Empty));
                sdsShowData_Main.InsertParameters.Add(new Parameter("RetireDate", DbType.Date, (eRetireDate_INS.Text.Trim() != "") ? eRetireDate_INS.Text.Trim() : String.Empty));
                sdsShowData_Main.InsertParameters.Add(new Parameter("TicketType", DbType.String, (eTicketType_INS.Text.Trim() != "") ? eTicketType_INS.Text.Trim() : String.Empty));
                sdsShowData_Main.InsertParameters.Add(new Parameter("Remark", DbType.String, (eRemark_INS.Text.Trim() != "") ? eRemark_INS.Text.Trim() : String.Empty));
                sdsShowData_Main.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                sdsShowData_Main.InsertParameters.Add(new Parameter("BuDate", DbType.Date, DateTime.Today.ToShortDateString()));
                sdsShowData_Main.Insert();

                fvRetireFreeTicketA.ChangeMode(FormViewMode.ReadOnly);
                OpenData();
                fvRetireFreeTicketA.DataBind();
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('資料有誤！員工編號欄位異常！')");
                Response.Write("</" + "Script>");
                eEmpNo_INS.Focus();
            }
        }

        /// <summary>
        /// 員工編號新增輸入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eEmpNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_Temp = "";
            string vWorkType = "";
            TextBox eEmpNo_INS = (TextBox)fvRetireFreeTicketA.FindControl("eEmpNo_INS");
            if (eEmpNo_INS != null)
            {
                if (eEmpNo_INS.Text.Trim() != "")
                {
                    vSQLStr_Temp = "select count(EmpNo) RCount from RetireFreeTicketA where EmpNo = '" + eEmpNo_INS.Text.Trim() + "' ";
                    if (Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "RCount")) == 0)
                    {
                        Label eEmpName_INS = (Label)fvRetireFreeTicketA.FindControl("eEmpName_INS");
                        Label eAssumeDate_INS = (Label)fvRetireFreeTicketA.FindControl("eAssumeDate_INS");
                        Label eRetireDate_INS = (Label)fvRetireFreeTicketA.FindControl("eRetireDate_INS");
                        Label eTicketType_INS = (Label)fvRetireFreeTicketA.FindControl("eTicketType_INS");
                        DropDownList ddlTicketType_INS = (DropDownList)fvRetireFreeTicketA.FindControl("ddlTicketType_INS");
                        vSQLStr_Temp = "select top 1 EmpNo, [Name], WorkType, AssumeDay, LeaveDay, " + Environment.NewLine +
                                       "       case when DateDiff(month, AssumeDay, LeaveDay) / 12 >= 25 then 3 " + Environment.NewLine +
                                       "            when DateDiff(month, AssumeDay, LeaveDay) / 12 between 15 and 24 then 2 " + Environment.NewLine +
                                       "            when DateDiff(month, AssumeDay, LeaveDay) / 12 between 10 and 14 then 1 " + Environment.NewLine +
                                       "            else 0 end TicketType" + Environment.NewLine +
                                       "  from Employee " + Environment.NewLine +
                                       " where EmpNo = '" + eEmpNo_INS.Text.Trim() + "' " + Environment.NewLine +
                                       " order by AssumeDay DESC";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            if (drTemp.HasRows)
                            {
                                while (drTemp.Read())
                                {
                                    vWorkType = drTemp["WorkType"].ToString().Trim();
                                    if (vWorkType == "退休")
                                    {
                                        eEmpName_INS.Text = drTemp["Name"].ToString().Trim();
                                        eAssumeDate_INS.Text = drTemp.GetDateTime(3).ToShortDateString();
                                        eRetireDate_INS.Text = drTemp.GetDateTime(4).ToShortDateString();
                                        eTicketType_INS.Text = drTemp["TicketType"].ToString().Trim();
                                        ddlTicketType_INS.SelectedIndex = Int32.Parse(drTemp["TicketType"].ToString().Trim());
                                    }
                                    else
                                    {
                                        Response.Write("<Script language='Javascript'>");
                                        Response.Write("alert('指定的員工狀態不正確！')");
                                        Response.Write("</" + "Script>");
                                        eEmpNo_INS.Focus();
                                    }
                                }
                            }
                            else
                            {
                                Response.Write("<Script language='Javascript'>");
                                Response.Write("alert('員工編號不存在！')");
                                Response.Write("</" + "Script>");
                                eEmpNo_INS.Focus();
                            }
                        }
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('指定的員工已有資料！')");
                        Response.Write("</" + "Script>");
                        eEmpNo_INS.Focus();
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('員工編號不可空白！')");
                    Response.Write("</" + "Script>");
                    eEmpNo_INS.Focus();
                }
            }
        }

        /// <summary>
        /// 優待類別新增變更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlTicketType_INS_TextChanged(object sender, EventArgs e)
        {
            DropDownList ddlTicketType_Temp = (DropDownList)fvRetireFreeTicketA.FindControl("ddlTicketType_INS");
            Label eTicketType_Temp = (Label)fvRetireFreeTicketA.FindControl("eTicketType_INS");
            if (ddlTicketType_Temp != null)
            {
                eTicketType_Temp.Text = ddlTicketType_Temp.SelectedValue.Trim();
                plShowData_Detail.Visible = (eTicketType_Temp.Text.Trim() != "0");
            }
        }

        /// <summary>
        /// 主檔資料刪除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDel_List_Click(object sender, EventArgs e)
        {
            Label eEmpNo_List = (Label)fvRetireFreeTicketA.FindControl("eEmpNo_List");
            if (eEmpNo_List != null)
            {
                string vSQLStr_Temp = "";
                string vModifyMode = "DELA";
                string vMaxIndex = "";
                string vIndexNo = "";
                int vNewIndex = 0;
                string vEmpNoItems = "";
                string vItems = "";
                string vIndexNoItemB = "";
                string vItemB = "";
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        //先備份到異動記錄
                        //備份明細
                        vSQLStr_Temp = "select IndexNo from RetireFreeTicketB where EmpNo = '" + eEmpNo_List.Text.Trim() + "' order by IndexNo ";
                        SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                        connTemp.Open();
                        SqlDataReader drTemp = cmdTemp.ExecuteReader();
                        while (drTemp.Read())
                        {
                            vIndexNo = drTemp["IndexNo"].ToString().Trim();
                            vSQLStr_Temp = "select max(ItemsB) maxIndex from RetireFreeTicketB_History where EmpNo = '" + eEmpNo_List.Text.Trim() + "' ";
                            vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "maxIndex");
                            vNewIndex = (vMaxIndex.Trim() != "") ? Int32.Parse(vMaxIndex) + 1 : 1;
                            vItemB = vNewIndex.ToString("D4");
                            vIndexNoItemB = eEmpNo_List.Text.Trim() + vItemB.Trim();
                            vSQLStr_Temp = "insert into RetireFreeTicketB_History " + Environment.NewLine +
                                           "       (IndexNoItemB, ModifyMode, IndexNo, ItemsB, EmpNo, Items, AssignDate, TicketYear, RemarkB, " + Environment.NewLine +
                                           "        BuMan, BuDate, ModifyMan, ModifyDate, CreateMan, CreateDate)" + Environment.NewLine +
                                           "select '" + vIndexNoItemB + "', '" + vModifyMode + "', IndexNo, '" + vItemB + "', EmpNo, Items, AssignDate, TicketYear, RemarkB, " + Environment.NewLine +
                                           "       BuMan, BuDate, ModifyMan, ModifyDate, '" + vLoginID + "', GetDate() " + Environment.NewLine +
                                           "  from RetireFreeTicketB " + Environment.NewLine +
                                           " where IndexNo = '" + vIndexNo + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        }

                        //備份主檔
                        vSQLStr_Temp = "select max(Items) maxIndex from RetireFreeTicketA_History where EmpNo = '" + eEmpNo_List.Text.Trim() + "' ";
                        vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "maxIndex");
                        vNewIndex = (vMaxIndex.Trim() != "") ? Int32.Parse(vMaxIndex) + 1 : 1;
                        vItems = vNewIndex.ToString("D4");
                        vEmpNoItems = eEmpNo_List.Text.Trim() + vItems.Trim();
                        vSQLStr_Temp = "insert into RetireFreeTicketA_History " + Environment.NewLine +
                                       "       (EmpNoItem, ModifyMode, EmpNo, Items, Name, AssumeDate, RetireDate, TicketType, Remark, " + Environment.NewLine +
                                       "        BuMan, BuDate, ModifyMan, ModifyDate, CreateMan, CreateDate) " + Environment.NewLine +
                                       "select '" + vEmpNoItems + "', '" + vModifyMode + "', EmpNo, '" + vItems + "', Name, AssumeDate, RetireDate, TicketType, Remark, " + Environment.NewLine +
                                       "       BuMan, BuDate, ModifyMan, ModifyDate, '" + vLoginID + "', GetDate() " + Environment.NewLine +
                                       "  from RetireFreeTicketA " + Environment.NewLine +
                                       " where EmpNo = '" + eEmpNo_List.Text.Trim() + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    }

                    //先刪明細
                    SqlDataSource sdsTemp = new SqlDataSource();
                    sdsTemp.ConnectionString = vConnStr;
                    sdsTemp.DeleteCommand = "delete RetireFreeTicketB where EmpNo = @EmpNo";
                    sdsTemp.DeleteParameters.Add(new Parameter("EmpNo", DbType.String, eEmpNo_List.Text.Trim()));
                    sdsTemp.Delete();

                    //再刪主檔
                    sdsShowData_Main.DeleteCommand = "delete RetireFreeTicketA where EmpNo = @EmpNo";
                    sdsShowData_Main.DeleteParameters.Clear();
                    sdsShowData_Main.DeleteParameters.Add(new Parameter("EmpNo", DbType.String, eEmpNo_List.Text.Trim()));
                    sdsShowData_Main.Delete();

                    gridShowData_Main.DataBind();
                    fvRetireFreeTicketA.DataBind();

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

        protected void fvRetireFreeTicketB_DataBound(object sender, EventArgs e)
        {
            string vDateURL = "";
            string vDateScript = "";

            switch (fvRetireFreeTicketB.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    TextBox eAssignDate_EditB = (TextBox)fvRetireFreeTicketB.FindControl("eAssignDate_EditB");
                    vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eAssignDate_EditB.ClientID;
                    vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eAssignDate_EditB.Attributes["onClick"] = vDateScript;

                    TextBox eTicketYear_EditB = (TextBox)fvRetireFreeTicketB.FindControl("eTicketYear_EditB");
                    eTicketYear_EditB.Text = (DateTime.Today.Year - 1911).ToString();
                    break;
                case FormViewMode.Insert:
                    TextBox eAssignDate_INSB = (TextBox)fvRetireFreeTicketB.FindControl("eAssignDate_INSB");
                    vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eAssignDate_INSB.ClientID;
                    vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eAssignDate_INSB.Attributes["onClick"] = vDateScript;
                    eAssignDate_INSB.Text = DateTime.Today.ToShortDateString();

                    TextBox eTicketYear_INSB = (TextBox)fvRetireFreeTicketB.FindControl("eTicketYear_INSB");
                    eTicketYear_INSB.Text = (DateTime.Today.Year - 1911).ToString();

                    Label eEmpNo_List = (Label)fvRetireFreeTicketA.FindControl("eEmpNo_List");
                    Label eEmpNo_INSB = (Label)fvRetireFreeTicketB.FindControl("eEmpNo_INSB");
                        eEmpNo_INSB.Text = eEmpNo_List.Text.Trim();

                        Label eEmpName_List = (Label)fvRetireFreeTicketA.FindControl("eEmpName_List");
                        Label eEmpName_INSB = (Label)fvRetireFreeTicketB.FindControl("eEmpName_INSB");
                        eEmpName_INSB.Text = eEmpName_List.Text.Trim();
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// 明細修改確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDetailOK_EditB_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eIndexNo_EditB = (Label)fvRetireFreeTicketB.FindControl("eIndexNo_EditB");
            Label eEmpNo_EditB = (Label)fvRetireFreeTicketB.FindControl("eEmpNo_EditB");
            string vModifyMode = "EDIT";
            string vMaxItems = "";
            int vNewItems = 0;
            string vItemB = "";
            string vIndexNoItemB = "";
            string vSQLStr_Temp = "";

            if (eIndexNo_EditB != null)
            {
                TextBox eAssignDate_EditB = (TextBox)fvRetireFreeTicketB.FindControl("eAssignDate_EditB");
                TextBox eTicketYear_EditB = (TextBox)fvRetireFreeTicketB.FindControl("eTicketYear_EditB");
                TextBox eRemarkB_EditB = (TextBox)fvRetireFreeTicketB.FindControl("eRemarkB_EditB");
                try
                {
                    //先備份明細
                    vSQLStr_Temp = "select max(ItemsB) MaxItems from RetireFreeTicketB_History where EmpNo = '" + eEmpNo_EditB.Text.Trim() + "' ";
                    vMaxItems = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems");
                    vNewItems = (vMaxItems.Trim() != "") ? Int32.Parse(vMaxItems) + 1 : 1;
                    vItemB = vNewItems.ToString("D4");
                    vIndexNoItemB = eEmpNo_EditB.Text.Trim() + vItemB.Trim();
                    vSQLStr_Temp = "insert into RetireFreeTicketB_History " + Environment.NewLine +
                                   "       (IndexNoItemB, ModifyMode, IndexNo, ItemsB, EmpNo, Items, AssignDate, TicketYear, RemarkB, " + Environment.NewLine +
                                   "        BuMan, BuDate, ModifyMan, ModifyDate, CreateMan, CreateDate)" + Environment.NewLine +
                                   "select '" + vIndexNoItemB + "', '" + vModifyMode + "', IndexNo, '" + vItemB + "', EmpNo, Items, AssignDate, TicketYear, RemarkB, " + Environment.NewLine +
                                   "       BuMan, BuDate, ModifyMan, ModifyDate, '" + vLoginID + "', GetDate() " + Environment.NewLine +
                                   "  from RetireFreeTicketB " + Environment.NewLine +
                                   " where IndexNo = '" + eIndexNo_EditB.Text.Trim() + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //修改明細資料
                    sdsShowData_Detail.UpdateCommand = "UPDATE RetireFreeTicketB " + Environment.NewLine +
                                                       "   SET AssignDate = @AssignDate, TicketYear = @TicketYear, RemarkB = @RemarkB, " + Environment.NewLine +
                                                       "       ModifyDate = @ModifyDate, ModifyMan = @ModifyMan " + Environment.NewLine +
                                                       " WHERE (IndexNo = @IndexNo)";
                    sdsShowData_Detail.UpdateParameters.Clear();
                    sdsShowData_Detail.UpdateParameters.Add(new Parameter("AssignDate", DbType.DateTime, (eAssignDate_EditB.Text.Trim() != "") ? eAssignDate_EditB.Text.Trim() : DateTime.Today.ToShortDateString()));
                    sdsShowData_Detail.UpdateParameters.Add(new Parameter("TicketYear", DbType.String, (eTicketYear_EditB.Text.Trim() != "") ? eTicketYear_EditB.Text.Trim() : String.Empty));
                    sdsShowData_Detail.UpdateParameters.Add(new Parameter("RemarkB", DbType.String, (eRemarkB_EditB.Text.Trim() != "") ? eRemarkB_EditB.Text.Trim() : String.Empty));
                    sdsShowData_Detail.UpdateParameters.Add(new Parameter("ModifyDate", DbType.DateTime, DateTime.Today.ToShortDateString()));
                    sdsShowData_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsShowData_Detail.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, eIndexNo_EditB.Text.Trim()));
                    sdsShowData_Detail.Update();

                    gridShowData_Detail.DataBind();
                    fvRetireFreeTicketB.ChangeMode(FormViewMode.ReadOnly);
                    fvRetireFreeTicketB.DataBind();
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
        /// 明細新增確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDetailOK_INSB_Click(object sender, EventArgs e)
        {
            Label eEmpNo_INSB = (Label)fvRetireFreeTicketB.FindControl("eEmpNo_INSB");
            if (eEmpNo_INSB != null)
            {
                TextBox eTicketYear_INSB = (TextBox)fvRetireFreeTicketB.FindControl("eTicketYear_INSB");
                TextBox eAssignDate_INSB = (TextBox)fvRetireFreeTicketB.FindControl("eAssignDate_INSB");
                if ((eTicketYear_INSB == null) || (eTicketYear_INSB.Text.Trim() == ""))
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('適用年度不可空白！')");
                    Response.Write("</" + "Script>");
                    eTicketYear_INSB.Focus();
                }
                else if ((eAssignDate_INSB == null) || (eAssignDate_INSB.Text.Trim() == ""))
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('申辦日期不可空白！')");
                    Response.Write("</" + "Script>");
                    eAssignDate_INSB.Focus();
                }
                else
                {
                    string vSQLStr_Temp = "";
                    string vIndexNo_Temp = "";
                    string vItems_Temp = "";
                    string vMaxItems = "";
                    int vNewItems = 0;

                    TextBox eRemarkB_INSB = (TextBox)fvRetireFreeTicketB.FindControl("eRemarkB_INSB");
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    try
                    {
                        vSQLStr_Temp = "select max(Items) MaxItems from RetireFreeTicketB where EmpNo = '" + eEmpNo_INSB.Text.Trim() + "' ";
                        vMaxItems = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems");
                        vNewItems = (vMaxItems.Trim() != "") ? Int32.Parse(vMaxItems) + 1 : 1;
                        vItems_Temp = vNewItems.ToString("D4");
                        vIndexNo_Temp = eEmpNo_INSB.Text.Trim() + vItems_Temp.Trim();

                        //寫入資料
                        sdsShowData_Detail.InsertCommand = "INSERT INTO RetireFreeTicketB " + Environment.NewLine +
                                                           "       (IndexNo, EmpNo, Items, AssignDate, TicketYear, RemarkB, BuMan, BuDate) " + Environment.NewLine +
                                                           "VALUES (@IndexNo, @EmpNo, @Items, @AssignDate, @TicketYear, @RemarkB, @BuMan, @BuDate)";
                        sdsShowData_Detail.InsertParameters.Clear();
                        sdsShowData_Detail.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo_Temp));
                        sdsShowData_Detail.InsertParameters.Add(new Parameter("EmpNo", DbType.String, eEmpNo_INSB.Text.Trim()));
                        sdsShowData_Detail.InsertParameters.Add(new Parameter("Items", DbType.String, vItems_Temp));
                        sdsShowData_Detail.InsertParameters.Add(new Parameter("AssignDate", DbType.DateTime, eAssignDate_INSB.Text.Trim()));
                        sdsShowData_Detail.InsertParameters.Add(new Parameter("TicketYear", DbType.String, eTicketYear_INSB.Text.Trim()));
                        sdsShowData_Detail.InsertParameters.Add(new Parameter("RemarkB", DbType.String, eRemarkB_INSB.Text.Trim()));
                        sdsShowData_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                        sdsShowData_Detail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, DateTime.Today.ToShortDateString()));
                        sdsShowData_Detail.Insert();

                        fvRetireFreeTicketB.ChangeMode(FormViewMode.ReadOnly);
                        gridShowData_Detail.DataBind();
                        fvRetireFreeTicketB.DataBind();

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
        }

        /// <summary>
        /// 明細資料刪除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDetailDel_ListB_Click(object sender, EventArgs e)
        {
            Label eIndexNo_ListB = (Label)fvRetireFreeTicketB.FindControl("eIndexNo_ListB");
            Label eEmpNo_ListB = (Label)fvRetireFreeTicketB.FindControl("eEmpNo_ListB");
            if (eIndexNo_ListB != null)
            {
                string vSQLStr_Temp = "";
                string vMaxItemB = "";
                int vNewItems = 0;
                string vIndexNoItemB = "";
                string vModifyMode = "DELB";
                string vItemB = "";

                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }

                try
                {
                    //先備份到異動檔
                    vSQLStr_Temp = "select max(ItemsB) MaxItems from RetireFreeTicketB_History where EmpNo = '" + eEmpNo_ListB.Text.Trim() + "' ";
                    vMaxItemB = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems");
                    vNewItems = (vMaxItemB.Trim() != "") ? Int32.Parse(vMaxItemB) + 1 : 1;
                    vItemB = vNewItems.ToString("D4");
                    vIndexNoItemB = eEmpNo_ListB.Text.Trim() + vItemB.Trim();
                    vSQLStr_Temp = "insert into RetireFreeTicketB_History " + Environment.NewLine +
                                   "       (IndexNoItemB, ModifyMode, IndexNo, ItemsB, EmpNo, Items, AssignDate, TicketYear, RemarkB, " + Environment.NewLine +
                                   "        BuMan, BuDate, ModifyMan, ModifyDate, CreateMan, CreateDate)" + Environment.NewLine +
                                   "select '" + vIndexNoItemB + "', '" + vModifyMode + "', IndexNo, '" + vItemB + "', EmpNo, Items, AssignDate, TicketYear, RemarkB, " + Environment.NewLine +
                                   "       BuMan, BuDate, ModifyMan, ModifyDate, '" + vLoginID + "', GetDate()) " + Environment.NewLine +
                                   "  from RetireFreeTicketB " + Environment.NewLine +
                                   " where IndexNo = '" + eIndexNo_ListB.Text.Trim() + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);

                    //刪除明細資料
                    sdsShowData_Detail.DeleteCommand = "delete RetireFreeTicketB where IndexNo = @IndexNo";
                    sdsShowData_Detail.DeleteParameters.Clear();
                    sdsShowData_Detail.DeleteParameters.Add(new Parameter("IndexNo", DbType.String, eIndexNo_ListB.Text.Trim()));
                    sdsShowData_Detail.Delete();
                    gridShowData_Detail.DataBind();
                    fvRetireFreeTicketB.DataBind();
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
    }
}