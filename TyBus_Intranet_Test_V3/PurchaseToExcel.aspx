<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="PurchaseToExcel.aspx.cs" Inherits="TyBus_Intranet_Test_V3.PurchaseToExcel" %>

<asp:Content ID="PurchaseToExcelForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">材料請購單轉出 EXCEL</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel" Width="100%">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbPurDate_Search" runat="server" CssClass="text-Right-Blue" Text="請購日期：" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eBuDateS_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eBuDateE_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:CheckBox ID="cbIncludeOil_Search" runat="server" CssClass="text-Left-Blue" Text="包括油料請購單" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:CheckBox ID="cbIncludeTire_Search" runat="server" CssClass="text-Left-Blue" Text="包括輪胎請購單" Width="100%" />
                </td>
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="離開" OnClick="bbClose_Click" Width="90%" />
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
