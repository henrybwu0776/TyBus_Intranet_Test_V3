<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CustomServiceListOut.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CustomServiceListOut" %>

<asp:Content ID="CustomServiceListOutForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">客服資料匯出</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuildDate" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eBuildDate_S" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eBuildDate_E" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col" colspan="8">
                    <asp:Label ID="lbInputNote_1" runat="server" CssClass="text-Left-Red" Font-Bold="true" Text="日期輸入方式：" />
                    <br />
                    <br />
                    <asp:Label ID="lbInputNote_2" runat="server" CssClass="text-Left-Red" Font-Bold="true" Text="1. 匯出所有資料：兩個日期欄位都保持空白" />
                    <br />
                    <br />
                    <asp:Label ID="lbInputNote_3" runat="server" CssClass="text-Left-Red" Font-Bold="true" Text="2. 匯出指定日期 (含) 之後的資料：只在前方的日期欄位輸入日期" />
                    <br />
                    <br />
                    <asp:Label ID="lbInputNote_4" runat="server" CssClass="text-Left-Red" Font-Bold="true" Text="3. 匯出指定日期 (含) 之前的資料：只在後方的日期欄位輸入日期" />
                    <br />
                    <br />
                    <asp:Label ID="lbInputNote_5" runat="server" CssClass="text-Left-Red" Font-Bold="true" Text="4. 匯出指定日期區間的資料：兩個日期欄位請都輸入日期" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel_2" runat="server" CssClass="button-Blue" Text="匯出分類統計" OnClick="bbExcel_2_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
            </tr>
        </table>
    </asp:Panel>
</asp:Content>
