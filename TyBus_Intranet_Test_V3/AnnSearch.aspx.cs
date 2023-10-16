using Amaterasu_Function;
using System;
using System.IO;
using System.Web;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class AnnSearch : System.Web.UI.Page
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
                    string vCaseDateURL = "InputDate.aspx?TextboxID=" + ePostDate_S_Search.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    ePostDate_S_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate.aspx?TextboxID=" + ePostDate_E_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    ePostDate_E_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate.aspx?TextboxID=" + eOpenDate_S_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eOpenDate_S_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate.aspx?TextboxID=" + eOpenDate_E_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eOpenDate_E_Search.Attributes["onClick"] = vCaseDateScript;

                    if (!IsPostBack)
                    {
                        ePostDate_S_Search.Text = PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "B");
                        ePostDate_E_Search.Text = PF.GetMonthLastDay(DateTime.Today, "B");
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

        private string GetSelStr()
        {
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eEmpNo_Search.Text.Trim() != "") ? "   and a.EmpNo '" + eEmpNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_PostDate = ((ePostDate_S_Search.Text.Trim() != "") && (ePostDate_E_Search.Text.Trim() != "")) ? "   and a.PostDate between '" + ePostDate_S_Search.Text.Trim() + "' and '" + ePostDate_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((ePostDate_S_Search.Text.Trim() != "") && (ePostDate_E_Search.Text.Trim() == "")) ? "   and a.PostDate = '" + ePostDate_S_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((ePostDate_S_Search.Text.Trim() == "") && (ePostDate_E_Search.Text.Trim() != "")) ? "   and a.PostDate = '" + ePostDate_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_OpenDate = ((eOpenDate_S_Search.Text.Trim() != "") && (eOpenDate_E_Search.Text.Trim() != "")) ? "   and a.EndDate >= '" + eOpenDate_S_Search.Text.Trim() + "' and a.StartDate <= '" + eOpenDate_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eOpenDate_S_Search.Text.Trim() != "") && (eOpenDate_E_Search.Text.Trim() == "")) ? "   and a.EndDate >= '" + eOpenDate_S_Search.Text.Trim() + "' and a.StartDate <= '" + eOpenDate_S_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eOpenDate_S_Search.Text.Trim() == "") && (eOpenDate_E_Search.Text.Trim() != "")) ? "   and a.EndDate >= '" + eOpenDate_E_Search.Text.Trim() + "' and a.StartDate <= '" + eOpenDate_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_PostTitle = (ePostTitle_Search.Text.Trim() != "") ? "   and a.PostTitle like '%" + ePostTitle_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_Remark = (eRemark_Search.Text.Trim() != "") ? "   and a.Remark like '%" + eRemark_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vSelectStr = "select AnnNo, DepNo, (select [Name] from Department where DepNo = a.DepNo) DepName, " + Environment.NewLine +
                                "       EmpNo, (select [Name] from Employee where EmpNo = a.EmpNo) EmpName, " + Environment.NewLine +
                                "       PostDate, StartDate, EndDate, PostTitle " + Environment.NewLine +
                                "  from AnnList a " + Environment.NewLine +
                                " where (1 = 1) " + Environment.NewLine +
                                vWStr_DepNo + vWStr_EmpNo + vWStr_OpenDate + vWStr_PostDate + vWStr_PostTitle + vWStr_Remark +
                                " order by PostDate DESC, DepNo, EmpNo";
            return vSelectStr;
        }

        private void OpenData()
        {
            string vSelectStr = GetSelStr();
            sdsAnnList_List.SelectCommand = vSelectStr;
            gridAnnList.DataBind();
        }

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_Search.Text.Trim();
            string vDepName_Temp = "";
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            if ((vDepNo_Temp.Trim() != "") && (vDepName_Temp.Trim() != ""))
            {
                eDepNo_Search.Text = vDepNo_Temp.Trim();
                eDepName_Search.Text = vDepName_Temp.Trim();
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('找不到您輸入的單位，請重新輸入！')");
                Response.Write("</" + "Script>");
                eDepNo_Search.Focus();
            }
        }

        protected void eEmpNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo_Temp = eEmpNo_Search.Text.Trim();
            string vEmpName_Temp = "";
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp + "' ";
            vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vEmpName_Temp == "")
            {
                vEmpName_Temp = vEmpNo_Temp;
                vSQLStr_Temp = "select EmpNo from Employee where [Name] = '" + vEmpName_Temp + "' ";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            if ((vEmpNo_Temp.Trim() != "") && (vEmpName_Temp.Trim() != ""))
            {
                eEmpNo_Search.Text = vEmpNo_Temp.Trim();
                eEmpName_Search.Text = vEmpName_Temp.Trim();
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('找不到您輸入的員工，請重新輸入！')");
                Response.Write("</" + "Script>");
                eEmpNo_Search.Focus();
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbDownLoad_List_Click(object sender, EventArgs e)
        {
            Label ePostFiles_List = (Label)fvAnnList_Detail.FindControl("ePostFiles_List");
            if (ePostFiles_List.Text.Trim() != "")
            {
                try
                {
                    string vDownLoadFiles = @"\\172.20.3.17\公用暫存區\26.各項公告資料\" + ePostFiles_List.Text.Trim();
                    string vFileName = ePostFiles_List.Text.Trim();
                    //============================================================================================================================
                    string vRecordNote = "下載檔案：公告資料查詢" + Environment.NewLine +
                                         "AnnSearch.aspx" + Environment.NewLine +
                                         "檔案名稱：" + vFileName;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    //============================================================================================================================
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
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
        }

        protected void fvAnnList_Detail_DataBound(object sender, EventArgs e)
        {
            switch (fvAnnList_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Button bbDownLoad_List = (Button)fvAnnList_Detail.FindControl("bbDownLoad_List");
                    Label ePostFiles_List = (Label)fvAnnList_Detail.FindControl("ePostFiles_List");
                    if (bbDownLoad_List != null)
                    {
                        bbDownLoad_List.Visible = (ePostFiles_List.Text.Trim() != "");
                    }
                    Label ePostOpen_List = (Label)fvAnnList_Detail.FindControl("ePostOpen_List");
                    CheckBox cbPostOpen_List = (CheckBox)fvAnnList_Detail.FindControl("cbPostOpen_List");
                    if (ePostOpen_List != null)
                    {
                        cbPostOpen_List.Checked = (ePostOpen_List.Text.Trim().ToUpper() == "V");
                    }
                    break;
                case FormViewMode.Edit:
                    break;
                case FormViewMode.Insert:
                    break;
            }
        }
    }
}