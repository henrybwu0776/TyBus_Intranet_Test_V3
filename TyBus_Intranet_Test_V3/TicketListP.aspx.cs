using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class TicketListP : System.Web.UI.Page
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
                        plPrint.Visible = false;
                        plShowData.Visible = true;
                    }
                    else
                    {
                        ListDataBind();
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

        private void ListDataBind()
        {
            sdsTicketListP.SelectCommand = GetSelectStr();
            gridTicketListP.DataBind();
        }

        private string GetSelectStr()
        {
            string vSelStr = "";
            string vWStr_TicketNo = ((eTicketNo_S.Text.Trim() != "") && (eTicketNo_E.Text.Trim() != "")) ? "   and tb.TicketNo between '" + eTicketNo_S.Text.Trim() + "' and '" + eTicketNo_E.Text.Trim() + "' " + Environment.NewLine :
                                    ((eTicketNo_S.Text.Trim() != "") && (eTicketNo_E.Text.Trim() == "")) ? "   and tb.TicketNo = '" + eTicketNo_S.Text.Trim() + "' " + Environment.NewLine :
                                    ((eTicketNo_S.Text.Trim() == "") && (eTicketNo_E.Text.Trim() != "")) ? "   and tb.TicketNo = '" + eTicketNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_PositionB = ((ePositionB_S.Text.Trim() != "") && (ePositionB_E.Text.Trim() != "")) ? "   and tc.PositionB between '" + ePositionB_S.Text.Trim() + "' and '" + ePositionB_E.Text.Trim() + "' " + Environment.NewLine :
                                     ((ePositionB_S.Text.Trim() != "") && (ePositionB_E.Text.Trim() == "")) ? "   and tc.PositionB = '" + ePositionB_S.Text.Trim() + "' " + Environment.NewLine :
                                     ((ePositionB_S.Text.Trim() == "") && (ePositionB_E.Text.Trim() != "")) ? "   and tc.PositionB = '" + ePositionB_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_LotNo = ((eLotNo_S.Text.Trim() != "") && (eLotNo_E.Text.Trim() != "")) ? "   and tb.LotNo between '" + eLotNo_S.Text.Trim() + "' and '" + eLotNo_E.Text.Trim() + "' " + Environment.NewLine :
                                 ((eLotNo_S.Text.Trim() != "") && (eLotNo_E.Text.Trim() == "")) ? "   and tb.LotNo = '" + eLotNo_S.Text.Trim() + "' " + Environment.NewLine :
                                 ((eLotNo_S.Text.Trim() == "") && (eLotNo_E.Text.Trim() != "")) ? "   and tb.LotNo = '" + eLotNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            //2019.07.17 修改語法....改用庫存檔為主體，如此一來只要沒有建庫存的票證資料就不會出現在這個列表裡 (原本是用票證B檔為主體)
            vSelStr = "select tb.TicketNo, (select TicketName from TickedsA where TicketNo = tb.TicketNo) TicketName, " + Environment.NewLine +
                      "       tc.PositionB, tb.LotNo, tc.NoStart, tc.NoEnd, isnull(tc.TicketCQty, 0) TicketQty " + Environment.NewLine +
                      "  from TickedsC tc left join TickedsB tb on tb.TicketNo = tc.ticketNo and tb.LotNo = tc.LotNo " + Environment.NewLine +
                      " where tb.TicketNo is not null " + Environment.NewLine + vWStr_TicketNo + vWStr_PositionB + vWStr_LotNo +
                      " order by tb.TicketNo, tb.LotNo, tc.PositionB ";
            return vSelStr;
        }

        protected void bbPreview_Click(object sender, EventArgs e)
        {
            ListDataBind();
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                string vSelStr = GetSelectStr();
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable();
                daPrint.Fill(dtPrint);

                ReportDataSource rdsPrint = new ReportDataSource("TicketListP", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\TicketListP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", PF.GetValue(vConnStr, "select [Name] from Custom where Types = 'O' and Code = 'A000'", "Name")));
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plShowData.Visible = false;

                string vTicketNoStr = ((eTicketNo_S.Text.Trim() != "") && (eTicketNo_E.Text.Trim() != "")) ? "從 " + eTicketNo_S.Text.Trim() + " 起至 " + eTicketNo_E.Text.Trim() + " 止" :
                                      ((eTicketNo_S.Text.Trim() != "") && (eTicketNo_E.Text.Trim() == "")) ? eTicketNo_S.Text.Trim() :
                                      ((eTicketNo_S.Text.Trim() == "") && (eTicketNo_E.Text.Trim() != "")) ? eTicketNo_E.Text.Trim() : "全部";
                string vPositionBStr = ((ePositionB_S.Text.Trim() != "") && (ePositionB_E.Text.Trim() != "")) ? "從 " + ePositionB_S.Text.Trim() + " 起至 " + ePositionB_E.Text.Trim() + " 止" :
                                       ((ePositionB_S.Text.Trim() != "") && (ePositionB_E.Text.Trim() == "")) ? ePositionB_S.Text.Trim() :
                                       ((ePositionB_S.Text.Trim() == "") && (ePositionB_E.Text.Trim() != "")) ? ePositionB_E.Text.Trim() : "全部";
                string vLotNoStr = ((eLotNo_S.Text.Trim() != "") && (eLotNo_E.Text.Trim() != "")) ? "從 " + eLotNo_S.Text.Trim() + " 起至 " + eLotNo_E.Text.Trim() + " 止" :
                                   ((eLotNo_S.Text.Trim() != "") && (eLotNo_E.Text.Trim() == "")) ? eLotNo_S.Text.Trim() :
                                   ((eLotNo_S.Text.Trim() == "") && (eLotNo_E.Text.Trim() != "")) ? eLotNo_E.Text.Trim() : "全部";
                string vRecordNote = "預覽報表_票證庫存統計表" + Environment.NewLine +
                                     "TicketListP.aspx" + Environment.NewLine +
                                     "票種代號：" + vTicketNoStr + Environment.NewLine +
                                     "庫位：" + vPositionBStr + Environment.NewLine +
                                     "批號：" + vLotNoStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vFileName = "票證庫存統計表";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                string vSelStr = GetSelectStr();
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    XSSFWorkbook wbExcel = new XSSFWorkbook();
                    //Excel 工作表
                    XSSFSheet wsExcel;
                    //設定標題欄位的格式
                    XSSFCellStyle csTitle = (XSSFCellStyle)wbExcel.CreateCellStyle();
                    csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                    csTitle.Alignment = HorizontalAlignment.Center; //水平置中

                    //設定字體格式
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
                    int vLinesNo = 0;

                    //新增一個工作表
                    wsExcel = (XSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i) == "TicketNo") ? "票種代號" :
                                      (drExcel.GetName(i) == "TicketName") ? "票種名稱" :
                                      (drExcel.GetName(i) == "PositionB") ? "庫位" :
                                      (drExcel.GetName(i) == "LotNo") ? "批號" :
                                      (drExcel.GetName(i) == "NoStart") ? "票號起" :
                                      (drExcel.GetName(i) == "NoEnd") ? "票號迄" :
                                      (drExcel.GetName(i) == "TicketQty") ? "數量" : drExcel.GetName(i);
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
                            if (drExcel.GetName(i) == "TicketQty")
                            {
                                if (drExcel[i].ToString() != "")
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(0);
                                }
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
                            string vTicketNoStr = ((eTicketNo_S.Text.Trim() != "") && (eTicketNo_E.Text.Trim() != "")) ? "從 " + eTicketNo_S.Text.Trim() + " 起至 " + eTicketNo_E.Text.Trim() + " 止" :
                                                  ((eTicketNo_S.Text.Trim() != "") && (eTicketNo_E.Text.Trim() == "")) ? eTicketNo_S.Text.Trim() :
                                                  ((eTicketNo_S.Text.Trim() == "") && (eTicketNo_E.Text.Trim() != "")) ? eTicketNo_E.Text.Trim() : "全部";
                            string vPositionBStr = ((ePositionB_S.Text.Trim() != "") && (ePositionB_E.Text.Trim() != "")) ? "從 " + ePositionB_S.Text.Trim() + " 起至 " + ePositionB_E.Text.Trim() + " 止" :
                                                   ((ePositionB_S.Text.Trim() != "") && (ePositionB_E.Text.Trim() == "")) ? ePositionB_S.Text.Trim() :
                                                   ((ePositionB_S.Text.Trim() == "") && (ePositionB_E.Text.Trim() != "")) ? ePositionB_E.Text.Trim() : "全部";
                            string vLotNoStr = ((eLotNo_S.Text.Trim() != "") && (eLotNo_E.Text.Trim() != "")) ? "從 " + eLotNo_S.Text.Trim() + " 起至 " + eLotNo_E.Text.Trim() + " 止" :
                                               ((eLotNo_S.Text.Trim() != "") && (eLotNo_E.Text.Trim() == "")) ? eLotNo_S.Text.Trim() :
                                               ((eLotNo_S.Text.Trim() == "") && (eLotNo_E.Text.Trim() != "")) ? eLotNo_E.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_票證庫存統計表" + Environment.NewLine +
                                                 "TicketListP.aspx" + Environment.NewLine +
                                                 "票種代號：" + vTicketNoStr + Environment.NewLine +
                                                 "庫位：" + vPositionBStr + Environment.NewLine +
                                                 "批號：" + vLotNoStr;
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
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            ListDataBind();
            plShowData.Visible = true;
            plPrint.Visible = false;
        }
    }
}