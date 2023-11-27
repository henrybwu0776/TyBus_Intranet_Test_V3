<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ExportStockList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ExportStockList" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">匯出材料庫存總表</a>
    </div>
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbPrintMode_Search" runat="server" CssClass="text-Right-Blue" Text="匯出條件" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:DropDownList ID="ePrintMode_Search" runat="server" CssClass="text-Left-Black" Width="95%">
                        <asp:ListItem Value="00" Text="全部庫存材料 (含零庫存)" />
                        <asp:ListItem Value="01" Text="只列出有庫存材料" />
                        <asp:ListItem Value="02" Text="只列出零庫存材料" />
                    </asp:DropDownList>
                </td>
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbExport" runat="server" CssClass="button-Blue" Text="匯出EXCEL" OnClick="bbExport_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" OnClick="bbExit_Click" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
            </tr>
        </table>
    </asp:Panel>
</asp:Content>
