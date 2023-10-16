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
    public partial class CarInfoExport : System.Web.UI.Page
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
        private string vSQLStr_GetList = "";

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
                        eDepNo_S.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "05") || (vLoginDepNo == "06") || (vLoginDepNo == "30"));
                        eDepNo_E.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "05") || (vLoginDepNo == "06") || (vLoginDepNo == "30"));
                        eDepNo_S.Text = ((Int32.Parse(vLoginDepNo) >= 11) && (Int32.Parse(vLoginDepNo) <= 30)) ? vLoginDepNo : "";
                        eCarYearCal.Text = DateTime.Today.ToString("d");
                        using (SqlConnection connGetList = new SqlConnection(vConnStr))
                        {
                            vSQLStr_GetList = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                              " union all " + Environment.NewLine +
                                              "select ClassNo, ClassTxt from DBDICB " + Environment.NewLine +
                                              " where FKey = '車輛資料作業    Car_infoA       TRAN_TYPE' " + Environment.NewLine +
                                              " order by ClassNo ";
                            SqlCommand cmdGetList = new SqlCommand(vSQLStr_GetList, connGetList);
                            connGetList.Open();
                            SqlDataReader drGetList = cmdGetList.ExecuteReader();
                            eCarState.Items.Clear();
                            while (drGetList.Read())
                            {
                                eCarState.Items.Add(new ListItem(drGetList["ClassTxt"].ToString().Trim(), drGetList["ClassNo"].ToString().Trim()));
                            }
                        }

                        using (SqlConnection connClass = new SqlConnection(vConnStr))
                        {
                            vSQLStr_GetList = "select ClassNo, ClassTxt from DBDICB " + Environment.NewLine +
                                              " where FKey = '車輛資料作業    Car_infoA       CLASS' " + Environment.NewLine +
                                              " order by ClassNo ";
                            SqlCommand cmdClass = new SqlCommand(vSQLStr_GetList, connClass);
                            connClass.Open();
                            SqlDataReader drClass = cmdClass.ExecuteReader();
                            while (drClass.Read())
                            {
                                eClass.Items.Add(new ListItem(drClass["ClassTxt"].ToString().Trim(), drClass["ClassNo"].ToString().Trim()));
                            }
                        }

                        using (SqlConnection connExceptional = new SqlConnection(vConnStr))
                        {
                            vSQLStr_GetList = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                              " union all " + Environment.NewLine +
                                              "select ClassNo, ClassTxt from DBDICB " + Environment.NewLine +
                                              " where FKey = '車輛資料作業    Car_infoA       EXCEPTIONAL' " + Environment.NewLine +
                                              " order by ClassNo ";
                            SqlCommand cmdExceptional = new SqlCommand(vSQLStr_GetList, connExceptional);
                            connExceptional.Open();
                            SqlDataReader drExceptional = cmdExceptional.ExecuteReader();
                            while (drExceptional.Read())
                            {
                                eExceptional.Items.Add(new ListItem(drExceptional["ClassTxt"].ToString().Trim(), drExceptional["ClassNo"].ToString().Trim()));
                            }
                        }
                    }
                    else
                    {
                        CarInfoDataBind();
                    }
                    string vBuildDateURL = "InputDate.aspx?TextboxID=" + eCarYearCal.ClientID;
                    string vBuildDateScript = "window.open('" + vBuildDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCarYearCal.Attributes["onClick"] = vBuildDateScript;
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
            DateTime vTempDate;
            string vCarYearCal = (DateTime.TryParse(eCarYearCal.Text.Trim(), out vTempDate)) ? vTempDate.ToShortDateString() : DateTime.Today.ToShortDateString();
            string vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "   and a.CompanyNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? "   and a.CompanyNo = '" + eDepNo_S.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? "   and a.CompanyNo = '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Point = ((ePoint_S.Text.Trim() != "") && (ePoint_E.Text.Trim() != "")) ? "   and a.Point between '" + ePoint_S.Text.Trim() + "' and '" + ePoint_E.Text.Trim() + "' " + Environment.NewLine :
                                 ((ePoint_S.Text.Trim() != "") && (ePoint_E.Text.Trim() == "")) ? "   and a.Point = '" + ePoint_S.Text.Trim() + "' " + Environment.NewLine :
                                 ((ePoint_S.Text.Trim() == "") && (ePoint_E.Text.Trim() != "")) ? "   and a.Point = '" + ePoint_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CarState = (eCarState.SelectedIndex > 0) ? "   and a.Tran_Type = '" + eCarState.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Class = (eClass.SelectedIndex > 0) ? "   and a.Class = '" + eClass.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Exceptional = (eExceptional.SelectedIndex > 0) ? "   and a.Exceptional = '" + eExceptional.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vSelStr = "select a.CompanyNo, (select [Name] from Department where DepNo = a.CompanyNo) CompanyName, " + Environment.NewLine +
                             "       a.Car_ID, a.Car_No, a.ProdDate, Year(a.ProdDate) ProdDate_Y, Month(a.ProdDate) ProdDate_M, "+Environment.NewLine+
                             "       Car_Class, (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CAR_CLASS' and ClassNo = a.Car_Class) Brand_C, Car_TypeID, " + Environment.NewLine +
                             "       (cast(Datediff(Month, a.ProdDate, '" + vCarYearCal + "') / 12 as varchar) + '年' + cast(DateDiff(Month, a.ProdDate, '" + vCarYearCal + "') % 12 as varchar) + '月') as CarYM," + Environment.NewLine +
                             "       a.SitQty, a.Tran_Type, (select ClassTxt from DBDICB where ClassNo = a.Tran_Type and FKey = '車輛資料作業    Car_infoA       TRAN_TYPE') Tran_Type_C, " + Environment.NewLine +
                             "       a.Point, (select ClassTxt from DBDICB where ClassNo = a.Point and FKey = '車輛資料作業    Car_infoA       POINT') Point_C " + Environment.NewLine +
                             "  from Car_InfoA a " + Environment.NewLine +
                             " where a.Tran_Type in ('1', '2') " + Environment.NewLine +
                             vWStr_DepNo + vWStr_Point + vWStr_CarState + vWStr_Class + vWStr_Exceptional +
                             " order by a.CompanyNo, a.ProdDate, a.Car_ID";
            return vSelStr;
        }

        private void CarInfoDataBind()
        {
            string vSelectStr = GetSelectStr();
            sdsCarInfo.SelectCommand = "";
            sdsCarInfo.SelectCommand = vSelectStr;
            gridCarInfo.DataBind();
        }

        protected void ddlPoint_S_SelectedIndexChanged(object sender, EventArgs e)
        {
            ePoint_S.Text = ddlPoint_S.SelectedValue;
        }

        protected void ddlPoint_E_SelectedIndexChanged(object sender, EventArgs e)
        {
            ePoint_E.Text = ddlPoint_E.SelectedValue;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            CarInfoDataBind();
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            //準備匯出EXCEL
            HSSFWorkbook wbExcel = new HSSFWorkbook();
            //Excel 工作表
            HSSFSheet wsExcel;

            //設定標題欄位的格式
            HSSFCellStyle csTitle = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中

            //設定標題欄位的字體格式
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

            string vFileName = "車輛資料匯出作業";
            int vLinesNo = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelectStr();
            SqlConnection connExcel = new SqlConnection(vConnStr);
            SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
            connExcel.Open();
            SqlDataReader drExcel = cmdExcel.ExecuteReader();
            if (drExcel.HasRows)
            {
                //查詢結果有資料的時候才執行
                DateTime vBuDate;
                //新增一個工作表
                wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                //寫入標題列
                vLinesNo = 0;
                wsExcel.CreateRow(vLinesNo);
                for (int i = 0; i < drExcel.FieldCount; i++)
                {
                    vHeaderText = (drExcel.GetName(i) == "CompanyNo") ? "部門編號" :
                                  (drExcel.GetName(i) == "CompanyName") ? "部門" :
                                  (drExcel.GetName(i) == "Car_ID") ? "牌照號碼" :
                                  (drExcel.GetName(i) == "Car_No") ? "車輛代號" :
                                  (drExcel.GetName(i) == "ProdDate") ? "出廠日期" :
                                  (drExcel.GetName(i) == "ProdDate_Y") ? "出廠年份" :
                                  (drExcel.GetName(i) == "ProdDate_M") ? "出廠月份" :
                                  (drExcel.GetName(i) == "CarYM") ? "車齡" :
                                  (drExcel.GetName(i) == "SitQty") ? "座位數" :
                                  (drExcel.GetName(i) == "Tran_Type") ? "狀態代碼" :
                                  (drExcel.GetName(i) == "Tran_Type_C") ? "狀態" :
                                  (drExcel.GetName(i) == "Point") ? "用途代碼" :
                                  (drExcel.GetName(i) == "Point_C") ? "用途" :
                                  (drExcel.GetName(i) == "Brand_C") ? "廠牌" :
                                  (drExcel.GetName(i) == "Car_Class") ? "廠牌代碼" :
                                  (drExcel.GetName(i) == "Car_TypeID") ? "型式" : drExcel.GetName(i);
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
                        if ((drExcel.GetName(i) == "ProdDate") && (drExcel[i].ToString() != ""))
                        {
                            string vTempStr = drExcel[i].ToString();
                            vBuDate = DateTime.Parse(drExcel[i].ToString());
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                        }
                        else if (((drExcel.GetName(i) == "ProdDate_Y") ||
                                  (drExcel.GetName(i) == "ProdDate_M") ||
                                  (drExcel.GetName(i) == "SitQty")) &&
                                 (drExcel[i].ToString() != ""))
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
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
                        string vRecordDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_S.Text.Trim() + "~" + eDepNo_E.Text.Trim() :
                                                 ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() :
                                                 ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() : "全部";
                        string vRecordPointStr = ((ePoint_S.Text.Trim() != "") && (ePoint_E.Text.Trim() != "")) ? ePoint_S.Text.Trim() + "~" + ePoint_E.Text.Trim() :
                                                 ((ePoint_S.Text.Trim() != "") && (ePoint_E.Text.Trim() == "")) ? ePoint_S.Text.Trim() :
                                                 ((ePoint_S.Text.Trim() == "") && (ePoint_E.Text.Trim() != "")) ? ePoint_E.Text.Trim() : "全部";
                        string vRecordNote = "匯出檔案_車輛資料匯出作業" + Environment.NewLine +
                                             "CarInfoExport.aspx" + Environment.NewLine +
                                             "站別：" + vRecordDepNoStr + Environment.NewLine +
                                             "用途：" + vRecordPointStr;
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

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}