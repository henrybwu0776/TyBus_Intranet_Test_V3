<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CalDepIncomeStatement.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CalDepIncomeStatement" %>

<asp:Content ID="CalDepIncomeStatementForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">客運業務損益表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalYM" runat="server" CssClass="text-Right-Blue" Text="計算年月：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eCalYear" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbCalYear" runat="server" CssClass="text-Left-Black" Text="年" />
                    <asp:TextBox ID="eCalMonth" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbCalMonth" runat="server" CssClass="text-Left-Black" Text="月" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBudgetYear" runat="server" CssClass="text-Right-Blue" Text="預算年度：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eBudgetYear" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="lbBudgetYearStr" runat="server" CssClass="text-Left-Black" Text="年" />
                </td>
                <td class="ColHeight ColWidth-8Col" colspan="2">
                    <asp:FileUpload ID="eFilePath" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbImportBudget" runat="server" CssClass="button-Black" Text="匯入預算資料" OnClick="bbImportBudget_Click" Width="90%" UseSubmitBehavior="False" OnClientClick="this.disabled=true;" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出業務損益表" OnClick="bbExcel_Click" Width="90%" />
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
                <td class="ColWidth-8Col" />
            </tr>
        </table>
    </asp:Panel>
</asp:Content>
