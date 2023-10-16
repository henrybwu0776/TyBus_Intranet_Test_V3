<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmpPayRangeList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmpPayRangeList" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="EmpPayRangeListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">員工薪資發放級距統計表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbPayYM_Search" runat="server" CssClass="text-Right-Blue" Text="發薪年月：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="ePayYear_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbPayYear_Search" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:DropDownList ID="ePayMonth_Search" runat="server" CssClass="text-Left-Black" Width="30%">
                        <asp:ListItem Text="1" Value="1" />
                        <asp:ListItem Text="2" Value="2" />
                        <asp:ListItem Text="3" Value="3" />
                        <asp:ListItem Text="4" Value="4" />
                        <asp:ListItem Text="5" Value="5" />
                        <asp:ListItem Text="6" Value="6" />
                        <asp:ListItem Text="7" Value="7" />
                        <asp:ListItem Text="8" Value="8" />
                        <asp:ListItem Text="9" Value="9" />
                        <asp:ListItem Text="10" Value="10" />
                        <asp:ListItem Text="11" Value="11" />
                        <asp:ListItem Text="12" Value="12" />
                    </asp:DropDownList>
                    <asp:Label ID="lbPayMonth_Search" runat="server" CssClass="text-Left-Black" Text=" 月" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbPayDur_Search" runat="server" CssClass="text-Right-Blue" Text="期別：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ePayDur_Search" runat="server" CssClass="text-Left-Black" Width="95%">
                        <asp:ListItem Selected="True" Text="1" Value="1" />
                        <asp:ListItem Text="2" Value="2" />
                        <asp:ListItem Text="3" Value="3" />
                    </asp:DropDownList>
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" OnClick="bbSearch_Click" Text="查詢" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Blue" OnClick="bbPrint_Click" Text="列印" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" OnClick="bbExcel_Click" Text="匯出EXCEL" Width="95%" />
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
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridEmpPatRangeList" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#DEDFDE" BorderStyle="None" BorderWidth="1px" CellPadding="4" DataSourceID="sdsEmpPayRangeList" ForeColor="Black" GridLines="Vertical" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:BoundField DataField="title" HeaderText="title" SortExpression="title" Visible="False" />
                <asp:BoundField DataField="Title_C" HeaderText="職稱" ReadOnly="True" SortExpression="Title_C" />
                <asp:BoundField DataField="RCount_1" HeaderText="少於24000" ReadOnly="True" SortExpression="RCount_1" />
                <asp:BoundField DataField="RCount_2" HeaderText="24001-40000" ReadOnly="True" SortExpression="RCount_2" />
                <asp:BoundField DataField="RCount_3" HeaderText="40001-50000" ReadOnly="True" SortExpression="RCount_3" />
                <asp:BoundField DataField="RCount_4" HeaderText="50001-60000" ReadOnly="True" SortExpression="RCount_4" />
                <asp:BoundField DataField="RCount_5" HeaderText="60001-70000" ReadOnly="True" SortExpression="RCount_5" />
                <asp:BoundField DataField="RCount_6" HeaderText="70001-80000" ReadOnly="True" SortExpression="RCount_6" />
                <asp:BoundField DataField="RCount_7" HeaderText="80001-90000" ReadOnly="True" SortExpression="RCount_7" />
                <asp:BoundField DataField="RCount_8" HeaderText="90001-100000" ReadOnly="True" SortExpression="RCount_8" />
                <asp:BoundField DataField="RCount_9" HeaderText="100001以上" ReadOnly="True" SortExpression="RCount_9" />
                <asp:BoundField DataField="TitleCount" HeaderText="人數合計" ReadOnly="True" SortExpression="TitleCount" />
            </Columns>
            <EmptyDataTemplate>
                <asp:Label ID="lbEmptyText" runat="server" CssClass="errorMessageText" Text="查無符合條件的資料！"></asp:Label>
            </EmptyDataTemplate>
            <FooterStyle BackColor="#CCCC99" />
            <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
            <RowStyle BackColor="#F7F7DE" />
            <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#FBFBF2" />
            <SortedAscendingHeaderStyle BackColor="#848384" />
            <SortedDescendingCellStyle BackColor="#EAEAD3" />
            <SortedDescendingHeaderStyle BackColor="#575357" />
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="plReport" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCloseReport" runat="server" CssClass="button-Red" OnClick="bbCloseReport_Click" Text="結束預覽" Width="120px" />
        <rsweb:ReportViewer ID="rvPrint" runat="server" BackColor="" ClientIDMode="AutoID" HighlightBackgroundColor=""
            InternalBorderColor="204, 204, 204" InternalBorderStyle="Solid" InternalBorderWidth="1px"
            LinkActiveColor="" LinkActiveHoverColor="" LinkDisabledColor=""
            PrimaryButtonBackgroundColor="" PrimaryButtonForegroundColor="" PrimaryButtonHoverBackgroundColor="" PrimaryButtonHoverForegroundColor=""
            SecondaryButtonBackgroundColor="" SecondaryButtonForegroundColor="" SecondaryButtonHoverBackgroundColor="" SecondaryButtonHoverForegroundColor=""
            SplitterBackColor=""
            ToolbarDividerColor="" ToolbarForegroundColor="" ToolbarForegroundDisabledColor=""
            ToolbarHoverBackgroundColor="" ToolbarHoverForegroundColor=""
            ToolBarItemBorderColor="" ToolBarItemBorderStyle="Solid" ToolBarItemBorderWidth="1px" ToolBarItemHoverBackColor=""
            ToolBarItemPressedBorderColor="51, 102, 153" ToolBarItemPressedBorderStyle="Solid" ToolBarItemPressedBorderWidth="1px"
            ToolBarItemPressedHoverBackColor="153, 187, 226" Height="600px" Width="100%" PageCountMode="Actual">
            <LocalReport ReportPath="Report\EmpPayRangeListP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsEmpPayRangeList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT title, CAST('' AS varchar) AS Title_C, CAST(0 AS int) AS TitleCount, CAST(0 AS int) AS RCount_1, CAST(0 AS int) AS RCount_2, CAST(0 AS int) AS RCount_3, CAST(0 AS int) AS RCount_4, CAST(0 AS int) AS RCount_5, CAST(0 AS int) AS RCount_6, CAST(0 AS int) AS RCount_7, CAST(0 AS int) AS RCount_8, CAST(0 AS int) AS RCount_9 FROM PAYREC AS p WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
