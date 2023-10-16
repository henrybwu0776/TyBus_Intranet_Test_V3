using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class EmpCardIDChange : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void bbGetNewID_Click(object sender, EventArgs e)
        {
            //2023.08.28 修改，原本 int 搭配 Int32 會出現長度不足的問題，改用 long 配合 Int64 進行轉換
            //int TempINT;
            long TempINT;
            //if ((eSourceID_Search.Text.Trim() != "") && (Int32.TryParse(eSourceID_Search.Text.Trim(), out TempINT)))
            if ((eSourceID_Search.Text.Trim() != "") && (Int64.TryParse(eSourceID_Search.Text.Trim(), out TempINT)))
            {
                string vSourceID = eSourceID_Search.Text.Trim();
                //string vNewID = Int32.Parse(vSourceID).ToString("X");
                string vNewID = Int64.Parse(vSourceID).ToString("X");
                string vOutputID = "";
                string[] vStrArray = new string[4];
                string vTempStr = "";
                for (int i = 0; i < vNewID.Length; i+=2)
                {
                    vTempStr = vNewID.Substring(i, 2);
                    vStrArray[i / 2] = vTempStr;
                }
                //eNewID_Search.Text = vNewID + "---";
                //vNewID += "---";
                for (int i = 3; i >= 0; i--)
                {
                    vOutputID += vStrArray[i];
                }
                eNewID_Search.Text = vOutputID;
            }

        }
    }
}