<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ImportAccount.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ImportAccount" %>

<asp:Content ID="ImportAccountForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">傳票匯入作業</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbAccDate_Search" runat="server" CssClass="text-Right-Blue" Text="傳票日期：" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eAccDate_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:CheckBox ID="eIgoneExcelDate_Search" runat="server" CssClass="text-Left-Blue" Text="一律使用傳票日期" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbAccType_Search" runat="server" CssClass="text-Right-Blue" Text="傳票類別：" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:DropDownList ID="ddlAccType_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbExcelName" runat="server" CssClass="text-Right-Blue" Text="來源檔案：" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="4">
                    <asp:FileUpload ID="fuExcel" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbImport" runat="server" CssClass="button-Blue" Text="開始轉入" OnClick="bbImport_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="離開" OnClick="bbClose_Click" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-7Col" />
                <td class="ColHeight ColWidth-7Col" />
                <td class="ColHeight ColWidth-7Col" />
                <td class="ColHeight ColWidth-7Col" />
                <td class="ColHeight ColWidth-7Col" />
                <td class="ColHeight ColWidth-7Col" />
                <td class="ColHeight ColWidth-7Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:TextBox ID="eErrorMessage" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Width="95%" Height="600px" />
        <asp:Button ID="bbCloseError" runat="server" CssClass="button-Red" Text="關閉訊息" Width="120px" OnClick="bbCloseError_Click" />
    </asp:Panel>
</asp:Content>
