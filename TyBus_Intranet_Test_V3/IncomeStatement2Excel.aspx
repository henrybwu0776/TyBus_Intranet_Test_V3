<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="IncomeStatement2Excel.aspx.cs" Inherits="TyBus_Intranet_Test_V3.IncomeStatement2Excel" %>

<asp:Content ID="IncomeStatement2ExcelForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">月損益表資料匯EXCEL</a>
    </div>
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalYM" runat="server" CssClass="text-Right-Blue" Text="匯出年月" Width="95%"></asp:Label>
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eCalYear" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="lbYearStr" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:TextBox ID="eCalMonth" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbMonthStr" runat="server" CssClass="text-Left-Black" Text=" 月" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExport" runat="server" CssClass="button-Blue" Text="匯出" OnClick="bbExport_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="95%" />
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
