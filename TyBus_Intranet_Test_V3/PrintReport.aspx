<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="PrintReport.aspx.cs" Inherits="TyBus_Intranet_Test_V3.PrintReport" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="PrintReportForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">報修單列印</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="6">
                    <asp:RadioButtonList ID="rbReportType" runat="server" CssClass="text-Left-Black" RepeatDirection="Horizontal" AutoPostBack="true"
                        OnSelectedIndexChanged="rbReportType_SelectedIndexChanged" Width="90%">
                        <asp:ListItem Value="00" Selected="True" Text="年度報表" />
                        <asp:ListItem Value="01" Text="月報表" />
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDep_Search" runat="server" CssClass="text-Right-Blue" Text="部門別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eDepNo_B_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_B_Search_TextChanged" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_E_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_E_Search_TextChanged" Width="40%" />
                    <br />
                    <asp:Label ID="eDepName_B_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="eDepName_E_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDay_Search" runat="server" CssClass="text-Right-Blue" Text="申請時間：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eDay_B_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDay_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="6">
                    <asp:RadioButtonList ID="rbGroupBy" runat="server" CssClass="text-Left-Black" RepeatDirection="Horizontal" AutoPostBack="true"
                        OnSelectedIndexChanged="rbGroupBy_SelectedIndexChanged" Width="90%">
                        <asp:ListItem Value="Report\SheetNoneGrouping.rdlc" Selected="True" Text="所有單據分開顯示" />
                        <asp:ListItem Value="Report\SheetGroupByBuMan.rdlc" Text="依申請人合併" />
                        <asp:ListItem Value="Report\SheetGroupByDepNo.rdlc" Text="依部門合併" />
                        <asp:ListItem Value="Report\SheetGroupByDate.rdlc" Text="依日期合併" />
                        <asp:ListItem Value="Report\SheetGroupByBuManAndDate.rdlc" Text="依申請人及日期合併" />
                        <asp:ListItem Value="Report\SheetGroupByDepNoAndDate.rdlc" Text="依部門及日期合併" />
                        <asp:ListItem Value="Report\SheetGroupByHandler.rdlc" Text="依處理人員及日期合併" />
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢報修單" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClear" runat="server" CssClass="button-Red" Text="清除重選" Width="90%" />
                </td>
                <td class="ColWidth-6Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Red" Text="預覽報表" OnClick="bbPrint_Click" Visible="false" Width="90%" />
                </td>
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridFixRequestList" runat="server" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" GridLines="None" OnPageIndexChanging="gridFixRequestList_PageIndexChanging" Width="100%">
            <FooterStyle BackColor="#C6C3C6" ForeColor="Black" />
            <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#E7E7FF" />
            <PagerStyle BackColor="#C6C3C6" ForeColor="Black" HorizontalAlign="Right" />
            <RowStyle BackColor="#DEDFDE" ForeColor="Black" />
            <SelectedRowStyle BackColor="#9471DE" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F1F1F1" />
            <SortedAscendingHeaderStyle BackColor="#594B9C" />
            <SortedDescendingCellStyle BackColor="#CAC9C9" />
            <SortedDescendingHeaderStyle BackColor="#33276A" />
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="plPrint" runat="server" Visible="false" CssClass="PrintPanel">
        <asp:Button ID="bbClosePrint" runat="server" CssClass="button-Red" Text="結束列印" OnClick="bbClosePrint_Click" Width="100px" />
        <rsweb:ReportViewer ID="rvPrint" runat="server" BackColor="" ClientIDMode="AutoID" HighlightBackgroundColor="" InternalBorderColor="204, 204, 204" InternalBorderStyle="Solid" InternalBorderWidth="1px" LinkActiveColor="" LinkActiveHoverColor="" LinkDisabledColor="" PrimaryButtonBackgroundColor="" PrimaryButtonForegroundColor="" PrimaryButtonHoverBackgroundColor="" PrimaryButtonHoverForegroundColor="" SecondaryButtonBackgroundColor="" SecondaryButtonForegroundColor="" SecondaryButtonHoverBackgroundColor="" SecondaryButtonHoverForegroundColor="" SplitterBackColor="" ToolbarDividerColor="" ToolbarForegroundColor="" ToolbarForegroundDisabledColor="" ToolbarHoverBackgroundColor="" ToolbarHoverForegroundColor="" ToolBarItemBorderColor="" ToolBarItemBorderStyle="Solid" ToolBarItemBorderWidth="1px" ToolBarItemHoverBackColor="" ToolBarItemPressedBorderColor="51, 102, 153" ToolBarItemPressedBorderStyle="Solid" ToolBarItemPressedBorderWidth="1px" ToolBarItemPressedHoverBackColor="153, 187, 226" Height="600px" Width="100%">
            <LocalReport ReportPath="Report\SheetGroupByBuMan.rdlc">
                <DataSources>
                    <rsweb:ReportDataSource DataSourceId="odsAllSheet" Name="FixRequestListP" />
                </DataSources>
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsFixRequestList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"></asp:SqlDataSource>
</asp:Content>
