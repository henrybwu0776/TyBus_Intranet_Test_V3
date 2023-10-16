using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class Lines952List : System.Web.UI.Page
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
        private DateTime vToday = DateTime.Today.AddDays(1);

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
                    string vCaseDateURL = "InputDate.aspx?TextboxID=" + eBuDate.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate.Attributes["onClick"] = vCaseDateScript;
                    if (!IsPostBack)
                    {
                        eBuDate.Text = vToday.Year.ToString("D4") + "/" + vToday.ToString("MM/dd");
                        eLinesNo.Text = "03560";
                    }
                    else
                    {
                        DataBinded();
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
            string vReturnStr = "";
            string vWStr_BuDate = "   and AssignNo in (select AssignNo from RunSheetA where BuDate = '" + eBuDate.Text.Trim() + "')" + Environment.NewLine;
            string vWStr_LinesNo = (eLinesNo.Text.Trim() != "") ? "   and LinesNo = '" + eLinesNo.Text.Trim() + "' " + Environment.NewLine : "";
            vReturnStr = "select (select BuDate from RunSheetA where AssignNo = b.AssignNo) DriveDate, " + Environment.NewLine +
                         "       (select[Name] from Department where DepNo = (select DepNo from RunSheetA where AssignNo = b.AssignNo)) DepName, " + Environment.NewLine +
                         "       (select Driver from RunSheetA where AssignNo = b.AssignNo) DriverNo, " + Environment.NewLine +
                         "       (select[Name] from Employee where EmpNo = (select Driver from RunSheetA where AssignNo = b.AssignNo)) DriverName, " + Environment.NewLine +
                         "       (select isnull(CallName, LineName) from Lines where LinesNo = b.LinesNo) LinesNo, " + Environment.NewLine +
                         "       Car_ID, ToTime, BackTime, Remark " + Environment.NewLine +
                         "  from RunSheetB b " + Environment.NewLine +
                         " where isnull(ReduceReason, '') = '' " + Environment.NewLine + vWStr_BuDate + vWStr_LinesNo +
                         " order by ToTime ";
            return vReturnStr;
        }

        private void DataBinded()
        {
            if (eBuDate.Text.Trim() != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSelectStr = GetSelStr();

                sdsDataShow.SelectCommand = "";
                sdsDataShow.SelectCommand = vSelectStr;
                gridDataShow.DataBind();
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請指定查詢日期！')");
                Response.Write("</" + "Script>");
                eBuDate.Focus();
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            DataBinded();
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            DateTime vBuDate;
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
            string vSelectStr = GetSelStr();
            string vFileName = "桃園客運路線" + vToday.ToString("yyyyMMdd") + "匯入新北動態系統班表資料";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
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
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("客運名稱：");
                    wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue("桃園客運");
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i) == "DriveDate") ? "發車日期" :
                                      (drExcel.GetName(i) == "DepName") ? "場站別" :
                                      (drExcel.GetName(i) == "DriverNo") ? "駕員編號" :
                                      (drExcel.GetName(i) == "DriverName") ? "駕員姓名" :
                                      (drExcel.GetName(i) == "LinesNo") ? "路線別" :
                                      (drExcel.GetName(i) == "Car_ID") ? "車號" :
                                      (drExcel.GetName(i) == "ToTime") ? "發車時間" :
                                      (drExcel.GetName(i) == "BackTime") ? "返站時間" :
                                      (drExcel.GetName(i) == "Remark") ? "備註" : drExcel.GetName(i);
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
                            if ((drExcel.GetName(i) == "DriveDate") && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
                            }
                            else if ((drExcel.GetName(i) == "DepName") && (drExcel[i].ToString() == "桃園公車站"))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue("桃園站");
                            }
                            else if (drExcel.GetName(i) == "Car_ID")
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString().Replace("-U-7", "-U7"));
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                            }
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
                            string vLinesNoStr = (eLinesNo.Text.Trim() != "") ? eLinesNo.Text.Trim() : "全部";
                            string vBuDateStr = (eBuDate.Text.Trim() != "") ? eBuDate.Text.Trim() : "";
                            string vRecordNote = "匯出檔案_路線匯出新北動態系統用資料" + Environment.NewLine +
                                                 "Lines952List.aspx" + Environment.NewLine +
                                                 "路線代號：" + vLinesNoStr + Environment.NewLine +
                                                 "行車日期：" + vBuDateStr;
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

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}