using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class WorkWeeklyReport : System.Web.UI.Page
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

                    //事件日期
                    string vCaseDateURL = "InputDate.aspx?TextboxID=" + eBuDate_S_Search.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate_S_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate.aspx?TextboxID=" + eBuDate_E_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate_E_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate.aspx?TextboxID=" + eUploadDate_S_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eUploadDate_S_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate.aspx?TextboxID=" + eUploadDate_E_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eUploadDate_E_Search.Attributes["onClick"] = vCaseDateScript;

                    if (!IsPostBack)
                    {
                        eDepNo_Search.Enabled = ((vLoginTitle == "080") ||
                                                 ((vLoginDepNo == "09") && (vLoginTitle == "140")) ||
                                                 ((vLoginDepNo == "09") && (vLoginTitle == "142")));
                        eDepNo_Search.Text = vLoginDepNo;
                        eDepName_Search.Text = vLoginDepName;
                        eEmpNo_Search.Enabled = ((vLoginTitle == "080") ||
                                                 ((vLoginDepNo == "09") && (vLoginTitle == "140")) ||
                                                 ((vLoginDepNo == "09") && (vLoginTitle == "142")));
                        eEmpNo_Search.Text = vLoginID;
                        eEmpName_Search.Text = vLoginName;
                        bbClear_Search.Visible = ((vLoginTitle == "080") ||
                                                  ((vLoginDepNo == "09") && (vLoginTitle == "140")) ||
                                                  ((vLoginDepNo == "09") && (vLoginTitle == "142")));
                        bbExcel_Search.Visible = ((vLoginTitle == "080") ||
                                                  ((vLoginDepNo == "09") && (vLoginTitle == "140")) ||
                                                  ((vLoginDepNo == "09") && (vLoginTitle == "142")));
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

        private string GetSelectStr()
        {
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eEmpNo_Search.Text.Trim() != "") ? "   and a.EmpNo = '" + eEmpNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_BuDate = ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() != "")) ? "   and a.BuDate between '" + eBuDate_S_Search.Text.Trim() + "' and '" + eBuDate_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                  ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() == "")) ? "   and a.BuDate = '" + eBuDate_S_Search.Text.Trim() + "' " + Environment.NewLine :
                                  ((eBuDate_S_Search.Text.Trim() == "") && (eBuDate_E_Search.Text.Trim() != "")) ? "   and a.BuDate = '" + eBuDate_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_UploadDate = ((eUploadDate_S_Search.Text.Trim() != "") && (eUploadDate_E_Search.Text.Trim() != "")) ? "   and a.UploadDate between '" + eUploadDate_S_Search.Text.Trim() + "' and '" + eUploadDate_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                      ((eUploadDate_S_Search.Text.Trim() != "") && (eUploadDate_E_Search.Text.Trim() == "")) ? "   and a.UploadDate = '" + eUploadDate_S_Search.Text.Trim() + "' " + Environment.NewLine :
                                      ((eUploadDate_S_Search.Text.Trim() == "") && (eUploadDate_E_Search.Text.Trim() != "")) ? "   and a.UploadDate = '" + eUploadDate_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vSelectStr = "SELECT IndexNo, DepNo, (SELECT [NAME] FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, " + Environment.NewLine +
                                "       EmpNo, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = a.EmpNo)) AS EmpName, " + Environment.NewLine +
                                "       Title, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '員工資料        EMPLOYEE        TITLE') AND (CLASSNO = a.Title)) AS Title_C, " + Environment.NewLine +
                                "       BuDate, BuMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuManName, " + Environment.NewLine +
                                "       UploadDate, UploadTime, FilePath, Remark, " + Environment.NewLine +
                                "       ModifyDate, ModifyMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = a.ModifyMan)) AS ModifyManName " + Environment.NewLine +
                                "  FROM WorkWeeklyReport AS a " + Environment.NewLine +
                                " WHERE ISNULL(IndexNo, '') <> '' " + Environment.NewLine +
                                vWStr_DepNo + vWStr_EmpNo + vWStr_BuDate + vWStr_UploadDate +
                                " ORDER BY DepNo, EmpNo, UploadDate DESC";
            return vSelectStr;
        }

        private void OpenData()
        {
            sdsWorkWeeklyReport_List.SelectCommand = "";
            sdsWorkWeeklyReport_List.SelectCommand = GetSelectStr();
            gridWorkWeeklyReport.DataBind();
        }

        private string GetIndexNo(string fIndexNo_Head)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vTempStr = "select MAX(IndexNo) as MaxNo from WorkWeeklyReport where IndexNo like '" + fIndexNo_Head + "%' ";
            string vOldMaxIndexNo = PF.GetValue(vConnStr, vTempStr, "MaxNo");
            Int32 vOldIndex = (vOldMaxIndexNo.Trim() != "") ? Int32.Parse(vOldMaxIndexNo.Replace(fIndexNo_Head, "")) : 0;
            string vNewIndexNo = fIndexNo_Head + (vOldIndex + 1).ToString("D4");
            return vNewIndexNo;
        }

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Search = eDepName_Search.Text.Trim();
            string vTempStr = "select [Name] from Department where DepNo = '" + vDepNo_Search.Trim() + "' ";
            string vDepName_Search = PF.GetValue(vConnStr, vTempStr, "Name");
            if (vDepName_Search == "")
            {
                vDepName_Search = vDepNo_Search.Trim();
                vTempStr = "select DepNo from Department where [Name] = '" + vDepName_Search.Trim() + "' ";
                vDepNo_Search = PF.GetValue(vConnStr, vTempStr, "DepNo");
            }
            eDepNo_Search.Text = vDepNo_Search.Trim();
            eDepName_Search.Text = vDepName_Search.Trim();
        }

        protected void eEmpNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo_Search = eEmpNo_Search.Text.Trim();
            string vTempStr = "select [Name] from Employee where EmpNo = '" + vEmpNo_Search + "' ";
            string vEmpName_Search = PF.GetValue(vConnStr, vTempStr, "Name");
            if (vEmpName_Search.Trim() == "")
            {
                vEmpName_Search = vEmpNo_Search.Trim();
                vTempStr = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_Search + "' order by AssumeDay DESC";
                vEmpNo_Search = PF.GetValue(vConnStr, vTempStr, "EmpNo");
            }
            eEmpNo_Search.Text = vEmpNo_Search.Trim();
            eEmpName_Search.Text = vEmpName_Search.Trim();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbClear_Search_Click(object sender, EventArgs e)
        {
            eDepNo_Search.Text = "";
            eEmpNo_Search.Text = "";
            eBuDate_S_Search.Text = "";
            eBuDate_E_Search.Text = "";
            eUploadDate_S_Search.Text = "";
            eUploadDate_E_Search.Text = "";
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvWorkWeeklyReport_DataBound(object sender, EventArgs e)
        {
            switch (fvWorkWeeklyReport.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Button bbDownload = (Button)fvWorkWeeklyReport.FindControl("bbDownload");
                    Button bbEdit = (Button)fvWorkWeeklyReport.FindControl("bbEdit");
                    Button bbDelete = (Button)fvWorkWeeklyReport.FindControl("bbDelete");
                    Button bbReupload = (Button)fvWorkWeeklyReport.FindControl("bbReupload");
                    FileUpload fuReupload = (FileUpload)fvWorkWeeklyReport.FindControl("fuReupload");
                    Label eEmpNo_List = (Label)fvWorkWeeklyReport.FindControl("eEmpNo_List");
                    if (bbDownload != null)
                    {
                        bbDownload.Visible = (vLoginTitle == "080");
                        bbEdit.Visible = (eEmpNo_List.Text.Trim() == vLoginID);
                        bbDelete.Visible = (((vLoginDepNo == "09") && (vLoginTitle == "140")) ||
                                            ((vLoginDepNo == "09") && (vLoginTitle == "142")));
                        bbReupload.Visible = (eEmpNo_List.Text.Trim() == vLoginID);
                        fuReupload.Visible = (eEmpNo_List.Text.Trim() == vLoginID);
                    }
                    break;
                case FormViewMode.Edit:
                    Label eEmpNo_Edit = (Label)fvWorkWeeklyReport.FindControl("eEmpNo_Edit");
                    Label eModifyDate_Edit = (Label)fvWorkWeeklyReport.FindControl("eModifyDate_Edit");
                    Label eModifyMan_Edit = (Label)fvWorkWeeklyReport.FindControl("eModifyMan_Edit");
                    Label eModifyManName_Edit = (Label)fvWorkWeeklyReport.FindControl("eModifyManName_Edit");

                    if (eEmpNo_Edit != null)
                    {
                        eModifyDate_Edit.Text = DateTime.Today.ToString("yyyy/MM/dd");
                        eModifyMan_Edit.Text = vLoginID;
                        eModifyManName_Edit.Text = vLoginName;
                    }
                    break;
                case FormViewMode.Insert:
                    Label eEmpNo_INS = (Label)fvWorkWeeklyReport.FindControl("eEmpNo_INS");
                    Label eEmpName_INS = (Label)fvWorkWeeklyReport.FindControl("eEmpName_INS");
                    Label eDepNo_INS = (Label)fvWorkWeeklyReport.FindControl("eDepNo_INS");
                    Label eDepName_INS = (Label)fvWorkWeeklyReport.FindControl("eDepName_INS");
                    Label eTitle_INS = (Label)fvWorkWeeklyReport.FindControl("eTitle_INS");
                    Label eTitle_C_INS = (Label)fvWorkWeeklyReport.FindControl("eTitle_C_INS");
                    Label eBuDate_INS = (Label)fvWorkWeeklyReport.FindControl("eBuDate_INS");
                    Label eBuMan_INS = (Label)fvWorkWeeklyReport.FindControl("eBuMan_INS");
                    Label eBuManName_INS = (Label)fvWorkWeeklyReport.FindControl("eBuManName_INS");
                    Label eFilePath_INS = (Label)fvWorkWeeklyReport.FindControl("eFilePath_INS");
                    TextBox eRemark_INS = (TextBox)fvWorkWeeklyReport.FindControl("eRemark_INS");

                    if (eEmpNo_INS != null)
                    {
                        eEmpNo_INS.Text = vLoginID;
                        eEmpName_INS.Text = vLoginName;
                        eDepNo_INS.Text = vLoginDepNo;
                        eDepName_INS.Text = vLoginDepName;
                        eTitle_INS.Text = vLoginTitle;
                        eTitle_C_INS.Text = vLoginTitleName;
                        eBuDate_INS.Text = DateTime.Today.ToString("yyyy/MM/dd");
                        eBuMan_INS.Text = vLoginID;
                        eBuManName_INS.Text = vLoginName;
                        eFilePath_INS.Text = "";
                        eRemark_INS.Text = "";
                    }
                    break;
                default:
                    break;
            }
        }

        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            Label eIndexNo_Edit = (Label)fvWorkWeeklyReport.FindControl("eIndexNo_Edit");
            TextBox eRemark_Edit = (TextBox)fvWorkWeeklyReport.FindControl("eRemark_Edit");
            if (eIndexNo_Edit != null)
            {
                string vIndexNo = eIndexNo_Edit.Text.Trim();
                string vTempStr = "update WorkWeeklyReport set Remark = @Remark, ModifyDate = @ModifyDate, ModifyMan = @ModifyMan where IndexNo = @IndexNo";
                try
                {
                    sdsWorkWeeklyReport_Detail.UpdateParameters.Clear();
                    sdsWorkWeeklyReport_Detail.UpdateCommand = vTempStr;
                    sdsWorkWeeklyReport_Detail.UpdateParameters.Add(new Parameter("Remark", DbType.String, (eRemark_Edit.Text.Trim() != "") ? eRemark_Edit.Text.Trim() : String.Empty));
                    sdsWorkWeeklyReport_Detail.UpdateParameters.Add(new Parameter("ModifyDate", DbType.DateTime, DateTime.Today.ToString("yyyy/MM/dd")));
                    sdsWorkWeeklyReport_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsWorkWeeklyReport_Detail.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                    sdsWorkWeeklyReport_Detail.Update();
                    fvWorkWeeklyReport.ChangeMode(FormViewMode.ReadOnly);
                    gridWorkWeeklyReport.DataBind();
                    fvWorkWeeklyReport.DataBind();
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
            string vIndexNo = GetIndexNo("W" + (DateTime.Today.Year - 1911).ToString("D3") + DateTime.Today.Month.ToString("D2") + vLoginID);
            string vInsertCommand = "INSERT INTO WorkWeeklyReport(IndexNo, DepNo, EmpNo, Title, BuDate, BuMan, UploadDate, UploadTime, FilePath, Remark) " + Environment.NewLine +
                                    "VALUES (@IndexNo, @DepNo, @EmpNo, @Title, @BuDate, @BuMan, @UploadDate, @UploadTime, @FilePath, @Remark)";
            string vUploadPath = @"\\172.20.3.17\各項上傳資料\工作周報\";

            Label eDepNo_INS = (Label)fvWorkWeeklyReport.FindControl("eDepNo_INS");
            Label eEmpNo_INS = (Label)fvWorkWeeklyReport.FindControl("eEmpNo_INS");
            Label eTitle_INS = (Label)fvWorkWeeklyReport.FindControl("eTitle_INS");
            Label eBuDate_INS = (Label)fvWorkWeeklyReport.FindControl("eBuDate_INS");
            Label eBuMan_INS = (Label)fvWorkWeeklyReport.FindControl("eBuMan_INS");
            FileUpload fuFilePath_INS = (FileUpload)fvWorkWeeklyReport.FindControl("fuFilePath_INS");
            TextBox eRemark_INS = (TextBox)fvWorkWeeklyReport.FindControl("eRemark_INS");
            if (fuFilePath_INS != null)
            {
                string vDepNo = eDepNo_INS.Text.Trim();
                string vEmpNo = eEmpNo_INS.Text.Trim();
                string vTitle = eTitle_INS.Text.Trim();
                string vBuDate = eBuDate_INS.Text.Trim();
                string vBuMan = eBuMan_INS.Text.Trim();
                string vRemark = eRemark_INS.Text.Trim();
                string vUploadDate = "";
                string vUploadTime = "";
                string vFilePath = "";
                string vFullPath = "";
                string vFileName = "";
                string vExtName = "";
                try
                {
                    if (fuFilePath_INS.FileName != "")
                    {
                        vUploadDate = DateTime.Today.ToString("yyyy/MM/dd");
                        vUploadTime = DateTime.Now.ToString("HH:mm");
                        vFileName = Path.GetFileNameWithoutExtension(fuFilePath_INS.FileName);
                        vExtName = Path.GetExtension(fuFilePath_INS.FileName);
                        vFilePath = vFileName + vIndexNo + vExtName;
                        vFullPath = vUploadPath + vFilePath;
                        fuFilePath_INS.SaveAs(vFullPath);
                    }

                    sdsWorkWeeklyReport_Detail.InsertParameters.Clear();
                    sdsWorkWeeklyReport_Detail.InsertCommand = vInsertCommand;
                    sdsWorkWeeklyReport_Detail.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                    sdsWorkWeeklyReport_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                    sdsWorkWeeklyReport_Detail.InsertParameters.Add(new Parameter("EmpNo", DbType.String, vEmpNo));
                    sdsWorkWeeklyReport_Detail.InsertParameters.Add(new Parameter("Title", DbType.String, vTitle));
                    sdsWorkWeeklyReport_Detail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, vBuDate));
                    sdsWorkWeeklyReport_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vBuMan));
                    sdsWorkWeeklyReport_Detail.InsertParameters.Add(new Parameter("UploadDate", DbType.Date, (vUploadDate.Trim() != "") ? vUploadDate : String.Empty));
                    sdsWorkWeeklyReport_Detail.InsertParameters.Add(new Parameter("UploadTime", DbType.String, (vUploadTime.Trim() != "") ? vUploadTime : String.Empty));
                    sdsWorkWeeklyReport_Detail.InsertParameters.Add(new Parameter("FilePath", DbType.String, (vFilePath.Trim() != "") ? vFilePath : String.Empty));
                    sdsWorkWeeklyReport_Detail.InsertParameters.Add(new Parameter("Remark", DbType.String, (vRemark.Trim() != "") ? vRemark : String.Empty));
                    sdsWorkWeeklyReport_Detail.Insert();
                    fvWorkWeeklyReport.ChangeMode(FormViewMode.ReadOnly);
                    gridWorkWeeklyReport.DataBind();
                    fvWorkWeeklyReport.DataBind();
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

        protected void bbDownload_Click(object sender, EventArgs e)
        {
            Label eFilePath_List = (Label)fvWorkWeeklyReport.FindControl("eFilePath_List");
            if ((eFilePath_List != null) && (eFilePath_List.Text.Trim() != ""))
            {
                string vDownLoadFiles = @"\\172.20.3.17\各項上傳資料\工作周報\" + eFilePath_List.Text.Trim();
                string vFileName = eFilePath_List.Text.Trim();
                FileStream fsTemp = new FileStream(vDownLoadFiles, FileMode.Open);
                byte[] bytes_Temp = new byte[(int)fsTemp.Length];
                fsTemp.Read(bytes_Temp, 0, bytes_Temp.Length);
                fsTemp.Close();
                Response.ContentType = "application/octet-stream";
                HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                string TourVerision = brObject.Type;
                if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                {
                    HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName, System.Text.Encoding.UTF8));
                }
                else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                {
                    // 設定強制下載標頭
                    Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName));
                }
                // 輸出檔案
                Response.BinaryWrite(bytes_Temp);
                Response.Flush();
                Response.End();
            }
        }

        protected void bbReupload_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eIndexNo_List = (Label)fvWorkWeeklyReport.FindControl("eIndexNo_List");
            FileUpload fuReupload = (FileUpload)fvWorkWeeklyReport.FindControl("fuReupload");
            if ((eIndexNo_List != null) && (eIndexNo_List.Text.Trim() != ""))
            {
                string vIndexNo = eIndexNo_List.Text.Trim();
                string vUploadPath = @"\\172.20.3.17\各項上傳資料\工作周報\";
                string vFileName = Path.GetFileNameWithoutExtension(fuReupload.FileName) + vIndexNo + Path.GetExtension(fuReupload.FileName);
                string vFullName = vUploadPath + vFileName;
                try
                {
                    if (File.Exists(vFullName))
                    {
                        File.Delete(vFullName);
                    }
                    string vUploadDate = DateTime.Today.ToString("yyyy/MM/dd");
                    string vUploadTime = DateTime.Now.ToString("HH:mm");
                    fuReupload.SaveAs(vFullName);
                    string vTempStr = "update WorkWeeklyReport set UploadDate = '" + vUploadDate + "', UploadTime = '" + vUploadTime + "', FilePath = '" + vFileName + "' where IndexNo = '" + vIndexNo + "' ";
                    PF.ExecSQL(vConnStr, vTempStr);
                    gridWorkWeeklyReport.DataBind();
                    fvWorkWeeklyReport.DataBind();
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

        protected void bbDelete_Click(object sender, EventArgs e)
        {
            Label eIndexNo_List = (Label)fvWorkWeeklyReport.FindControl("eIndexNo_List");
            Label eFilePath_List = (Label)fvWorkWeeklyReport.FindControl("eFilePath_List");
            if (eIndexNo_List != null)
            {
                string vIndexNo_List = eIndexNo_List.Text.Trim();
                string vUploadPath = @"\\172.20.3.17\各項上傳資料\工作周報\";
                string vFilePath_List = eFilePath_List.Text.Trim();
                string vFullPath = vUploadPath + vFilePath_List;
                string vDeleteCommand = "delete WorkWeeklyReport where IndexNo = '" + vIndexNo_List + "' ";
                try
                {
                    sdsWorkWeeklyReport_Detail.DeleteParameters.Clear();
                    sdsWorkWeeklyReport_Detail.DeleteCommand = vDeleteCommand;
                    sdsWorkWeeklyReport_Detail.DeleteParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo_List));
                    sdsWorkWeeklyReport_Detail.Delete();
                    if (File.Exists(vFullPath))
                    {
                        File.Delete(vFullPath);
                    }
                    gridWorkWeeklyReport.DataBind();
                    fvWorkWeeklyReport.DataBind();
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

        protected void bbExcel_Search_Click(object sender, EventArgs e)
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
            string vSelStr = "";
            string vFileName = "工作週報清冊";
            string vSheetName = "";
            DateTime vDate;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            //產生工作表
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                vSelStr = GetSelectStr();
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    vSheetName = vFileName.Trim();
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSheetName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "INDEXNO") ? "序號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNO") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "單位" :
                                      (drExcel.GetName(i).ToUpper() == "EMPNO") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "EMPNAME") ? "姓名" :
                                      (drExcel.GetName(i).ToUpper() == "TITLE") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "TITLE_C") ? "職稱" :
                                      (drExcel.GetName(i).ToUpper() == "BUDATE") ? "建立日期" :
                                      (drExcel.GetName(i).ToUpper() == "UPLOADDATE") ? "上傳日期" :
                                      (drExcel.GetName(i).ToUpper() == "UPLOADTIME") ? "上傳時間" :
                                      (drExcel.GetName(i).ToUpper() == "BUMAN") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "BUMANNAME") ? "開單人員" :
                                      (drExcel.GetName(i).ToUpper() == "FILEPATH") ? "附件檔案" :
                                      (drExcel.GetName(i).ToUpper() == "REMARK") ? "備註" :
                                      (drExcel.GetName(i).ToUpper() == "MODIFYMAN") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "MODIFYMANNAME") ? "異動人員" :
                                      (drExcel.GetName(i).ToUpper() == "MODIFYDATE") ? "異動日期" : drExcel.GetName(i);
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    while (drExcel.Read())
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            if (((drExcel.GetName(i).ToUpper() == "BUDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "UPLOADDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "MODIFYDATE")) &&
                                (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vDate.Year.ToString("D4") + "/" + vDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                        }
                    }
                }
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
                Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                Response.Write("</" + "Script>");
            }
        }
    }
}