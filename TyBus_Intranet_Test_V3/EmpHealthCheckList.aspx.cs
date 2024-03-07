using Amaterasu_Function;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.UI;

namespace TyBus_Intranet_Test_V3
{
    public partial class EmpHealthCheckList : Page
    {
        PublicFunction PF = new PublicFunction(); //加入公用程式碼參考
        private string vConnStr = "";
        private string vLoginID = "";
        private string vLoginName = "";
        private string vLoginDepNo = "";
        private string vLoginDepName = "";
        private string vLoginTitle = "";
        private string vLoginTitleName = "";
        private string vLoginEmpType = "";
        private string vComputerName = ""; //2021.09.27 新增

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
                    string vCaseDateURL = "InputDate.aspx?TextboxID=" + eCalBaseDate.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCalBaseDate.Attributes["onClick"] = vCaseDateScript;
                    if (!IsPostBack)
                    {
                        eCalBaseDate.Text = PF.GetMonthFirstDay(DateTime.Today, "B");
                        lbErrorMSG.Visible = false;
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

        private DataTable OpenData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DateTime dtBaseDate;
            DataTable dtResult = new DataTable();
            string vCalBaseDate = DateTime.TryParse(eCalBaseDate.Text.Trim(), out dtBaseDate) ? dtBaseDate.ToShortDateString() : PF.GetMonthFirstDay(DateTime.Today, "B");
            string vSelectStr = "select d.[Name] DepName, t.ClassTxt Title_C, e.EmpNo, e.[Name] EmpName, convert(varchar(10), e.Birthday, 23) Birthday, " + Environment.NewLine +
                                "       case when e.Citzen = '1' then '台灣' else '外籍' end Citzen_C, e.IDCardNo, e.Gender, cast('是' as varchar) IsEmployee, " + Environment.NewLine +
                                "       convert(varchar(10), e.Assumeday, 23) Assumeday, " + Environment.NewLine +
                                "       case when isnull(e.Cellphone, '') <> '' then e.CellPhone " + Environment.NewLine +
                                "            when isnull(e.TelNo1, '') <> '' then e.TelNo1 " + Environment.NewLine +
                                "            when isnull(e.TelNo2, '') <> '' then e.TelNo2 else null end Telephone " + Environment.NewLine +
                                "  from Employee e left join Department d on d.DepNo = e.DepNo " + Environment.NewLine +
                                "                  left join DBDICB t on t.ClassNo = e.Title and t.FKey = '員工資料        EMPLOYEE        TITLE' " + Environment.NewLine +
                                " where Assumeday <= @Assumeday and isnull(Leaveday, '') = '' " + Environment.NewLine +
                                "   and e.DepNo > '00' " + Environment.NewLine +
                                "   and isnull(e.[Type], '') <> '' " + Environment.NewLine +
                                " order by e.DepNo, e.EmpNo";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daTemp = new SqlDataAdapter(vSelectStr, connTemp);
                daTemp.SelectCommand.Parameters.Clear();
                daTemp.SelectCommand.Parameters.Add(new SqlParameter("Assumeday", vCalBaseDate));
                connTemp.Open();
                daTemp.Fill(dtResult);
            }
            return dtResult;
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            lbErrorMSG.Visible = false;
            XSSFWorkbook wbExcel = new XSSFWorkbook();
            //Excel 工作表
            XSSFSheet wsExcel = new XSSFSheet();

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

            string vFileName = "年度健康檢查名冊";
            int vLinesNo = 0;

            DataTable dtExcel = OpenData();
            if (dtExcel.Rows.Count>0)
            {
                string vSheetName = "";
                for (int iCount = 0; iCount < dtExcel.Rows.Count; iCount++)
                {
                    if (vSheetName != dtExcel.Rows[iCount]["DepName"].ToString().Trim())
                    {
                        //新增一個工作表
                        vLinesNo = 0;
                        wsExcel = (XSSFSheet)wbExcel.CreateSheet(dtExcel.Rows[iCount]["DepName"].ToString().Trim());
                        vSheetName = dtExcel.Rows[iCount]["DepName"].ToString().Trim();
                        //寫入標題列
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("部門");
                        wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue("職稱");
                        wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("員工代號");
                        wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("姓名");
                        wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("出生日期");
                        wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("國籍");
                        wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("身份證");
                        wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("性別");
                        wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("是否為員工");
                        wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("受雇日期");
                        wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("電話");
                        wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellValue("備註");
                        wsExcel.GetRow(vLinesNo).GetCell(11).CellStyle = csTitle;
                        wsExcel.GetRow(vLinesNo).CreateCell(12).SetCellValue("預訂健檢日期");
                        wsExcel.GetRow(vLinesNo).GetCell(12).CellStyle = csTitle;
                    }
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(dtExcel.Rows[iCount]["DepName"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue(dtExcel.Rows[iCount]["Title_C"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue(dtExcel.Rows[iCount]["EmpNo"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue(dtExcel.Rows[iCount]["EmpName"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue(dtExcel.Rows[iCount]["Birthday"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue(dtExcel.Rows[iCount]["Citzen_C"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue(dtExcel.Rows[iCount]["IDCardNo"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue(dtExcel.Rows[iCount]["Gender"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue(dtExcel.Rows[iCount]["IsEmployee"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue(dtExcel.Rows[iCount]["Assumeday"].ToString().Trim());
                    wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData;
                    string vTempStr = dtExcel.Rows[iCount]["Telephone"].ToString().Trim().Replace("-", "");
                    string vTelePhone = ((vTempStr != "") && (vTempStr.Length >= 10)) ? vTempStr.Substring(0, 4) + "-" + vTempStr.Substring(4, 3) + "-" + vTempStr.Substring(7) :
                                        ((vTempStr != "") && (vTempStr.Length >= 7)) ? vTempStr.Substring(0, 2) + "-" + vTempStr.Substring(2) : vTempStr;
                    wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue(vTelePhone);
                    wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData;
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
                        string vRecordNote = "匯出檔案_年度健康檢查名冊" + Environment.NewLine +
                                             "EmpHealthCheckList.aspx";
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                        //===========================================================================================================
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
                    lbErrorMSG.Text = eMessage.ToString();
                    lbErrorMSG.Visible = true;
                }
            }
            else
            {
                lbErrorMSG.Text = "查無符合條件的資料";
                lbErrorMSG.Visible = true;
            }
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}