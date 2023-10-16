<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmpCardIDChange.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmpCardIDChange" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">人員卡內碼轉換</a>
    </div>
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbSourceID_Search" runat="server" CssClass="text-Right-Blue" Text="十碼數字內碼：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eSourceID_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Button ID="bbGetNewID" runat="server" CssClass="button-Red" Text="轉換—＞" OnClick="bbGetNewID_Click" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbNewID_Search" runat="server" CssClass="text-Right-Blue" Text="中保系統用 ID：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eNewID_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
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
</asp:Content>
