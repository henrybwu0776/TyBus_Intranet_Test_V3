<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DocFirstWordPermission.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DocFirstWordPermission" %>

<asp:Content ID="DocFirstWordPermissionForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">公文字號權限設定</a>
    </div>
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbEmpNo_Search" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eEmpNo_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="eEmpName_Search" runat="server" CssClass="text-Left-Black" Width="50%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDocFirstWord_Search" runat="server" CssClass="text-Right-Blue" Text="公文字號：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:DropDownList ID="ddlDocFirstWord_Search" runat="server" CssClass="text-Left-Black" OnSelectedIndexChanged="ddlDocFirstWord_Search_SelectedIndexChanged" Width="95%" />
                    <asp:Label ID="eDocFirstWord_Search" runat="server" Visible="false" />
                 </td>
                <td class="ColHeight ColWidth-8Col" colspan="2">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Blue" Text="查詢" OnClick="bbSearch_Click" Width="120px" />
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="結束" OnClick="bbExit_Click" Width="120px" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
            </tr>
        </table>
    </asp:Panel>
</asp:Content>
