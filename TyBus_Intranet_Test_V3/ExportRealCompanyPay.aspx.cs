using Amaterasu_Function;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class ExportRealCompanyPay : System.Web.UI.Page
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
        private DateTime vFirstDate;
        private DateTime vLastDate;

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
                DateTime vToday = DateTime.Today;

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        int vYear = ((vToday.Month - 1) <= 0) ? vToday.Year - 1 : vToday.Year;
                        int vMonth = ((vToday.Month - 1) <= 0) ? vToday.Month + 11 : vToday.Month - 1;
                        eCalYM.Text = (eCalYM.Text.Trim() == "") ? (vYear.ToString("D4") + "/" + vMonth.ToString("D2")) : eCalYM.Text.Trim();
                    }
                    bbExcel.Visible = (gridShowData.Rows.Count > 0);
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
            vFirstDate = (eCalYM.Text.Trim() != "") ? DateTime.Parse(eCalYM.Text.Trim() + "/01") : DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today, "B"));
            vLastDate = DateTime.Parse(PF.GetMonthLastDay(vFirstDate, "B"));
            string vStartDateStr = vFirstDate.Year.ToString() + "/" + vFirstDate.ToString("MM/dd");
            string vEndDateStr = vLastDate.Year.ToString() + "/" + vLastDate.ToString("MM/dd");
            string vSelectStr = "SELECT depno as DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.depno)) AS DepName, " + Environment.NewLine +
                                "       empno as EmpNo, [name] as [Name], paydate as PayDate, " + Environment.NewLine +
                                "       cashnum71 as CashNum71, cashnum72 as CashNum72, cashnum74 as CashNum74 " + Environment.NewLine +
                                "  FROM PAYREC AS e " + Environment.NewLine +
                                " WHERE PayDate between '" + vStartDateStr + "' and '" + vEndDateStr + "' " + Environment.NewLine +
                                "   AND PayDur = '1' " + Environment.NewLine +
                                " ORDER BY DepNo, EmpNo";
            return vSelectStr;
        }

        private void OpenData()
        {
            string vSelStr = GetSelectStr();
            sdsShowData.SelectCommand = "";
            sdsShowData.SelectCommand = vSelStr;
            gridShowData.DataBind();
            bbExcel.Visible = (gridShowData.Rows.Count > 0);
        }

        protected void bbOK_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            //string vHeadText = "";
            string vGetDate = "";
            DateTime vTempDate;
            //Excel 檔案
            XSSFWorkbook wbExcel = new XSSFWorkbook();
            //Excel 工作表
            XSSFSheet wsExcel = (XSSFSheet)wbExcel.CreateSheet("公司提撥資料");
            //設定標題欄位的格式
            XSSFCellStyle csTitle = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中
                                                            //設定字體格式
            XSSFFont fontTitle = (XSSFFont)wbExcel.CreateFont();
            //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
            fontTitle.IsBold = true;
            fontTitle.FontHeightInPoints = 20; //字體大小
            csTitle.SetFont(fontTitle);

            //建立標題列
            wsExcel.CreateRow(0);
            for (int i = 0; i < gridShowData.Columns.Count; i++)
            {
                wsExcel.GetRow(0).CreateCell(i).SetCellValue(gridShowData.Columns[i].HeaderText.Trim());
            }
            for (int j = 0; j < gridShowData.Rows.Count; j++)
            {
                wsExcel.CreateRow(j + 1);
                for (int k = 0; k < gridShowData.Columns.Count; k++)
                {
                    if ((gridShowData.Columns[k].HeaderText == "勞保公司負擔") ||
                        (gridShowData.Columns[k].HeaderText == "健保公司負擔") ||
                        (gridShowData.Columns[k].HeaderText == "勞退公司提撥"))
                    {
                        //這幾個欄位是數值欄位
                        wsExcel.GetRow(j + 1).CreateCell(k);
                        wsExcel.GetRow(j + 1).GetCell(k).SetCellType(CellType.Numeric);
                        wsExcel.GetRow(j + 1).GetCell(k).SetCellValue(double.Parse(gridShowData.Rows[j].Cells[k].Text.Trim()));
                    }
                    else
                    {
                        wsExcel.GetRow(j + 1).CreateCell(k);
                        wsExcel.GetRow(j + 1).GetCell(k).SetCellType(CellType.String);
                        if (gridShowData.Columns[k].HeaderText == "計薪年月")
                        {
                            vTempDate = DateTime.Parse(gridShowData.Rows[j].Cells[k].Text.Trim());
                            vGetDate = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                            wsExcel.GetRow(j + 1).GetCell(k).SetCellValue(vGetDate);
                        }
                        else
                        {
                            wsExcel.GetRow(j + 1).GetCell(k).SetCellValue(gridShowData.Rows[j].Cells[k].Text.Trim());
                        }
                    }
                }
            }
            try
            {
                /*
                MemoryStream msTarget = new MemoryStream();
                wbExcel.Write(msTarget);
                // 設定強制下載標頭
                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=匯出公司提撥金額.xlsx"));
                // 輸出檔案
                Response.BinaryWrite(msTarget.ToArray());
                msTarget.Close();

                Response.End(); //*/
                string vFileName = "匯出公司提撥金額";
                var msTarget = new NPOIMemoryStream();
                msTarget.AllowClose = false;
                wbExcel.Write(msTarget);
                msTarget.Flush();
                msTarget.Seek(0, SeekOrigin.Begin);
                msTarget.AllowClose = true;

                if (msTarget.Length > 0)
                {
                    string vCalYMStr = eCalYM.Text.Trim();
                    string vRecordNote = "匯出檔案_匯出公司提撥金額" + Environment.NewLine +
                                         "ExportRealCompanyPay.aspx" + Environment.NewLine +
                                         "查核年月：" + vCalYMStr;
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