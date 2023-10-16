<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="TimeTableToGov.aspx.cs" Inherits="TyBus_Intranet_Test_V3.TimeTableToGov" %>
<asp:Content ID="TimeTableToGovForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">路線班次匯出公總用資料</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="主責單位：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuDate_Search" runat="server" CssClass="text-Right-Blue" Text="行車日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eBuDate_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" OnClick="bbExcel_Click" Text="匯出 EXCEL" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" OnClick="bbClose_Click" Text="結束" Width="95%" />
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
                <td class="ColWidth-8Col" />
            </tr>
        </table>
        </asp:Panel>
</asp:Content>
