<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="AnecdoteCaseP.aspx.cs" Inherits="TyBus_Intranet_Test_V3.AnecdoteCaseP" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="AnecdoteCasePForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">肇事案件統計表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="8">
                    <asp:RadioButtonList ID="rbSelectPrintMode" runat="server" CssClass="text-Left-Black" Width="90%" RepeatColumns="3">
                        <asp:ListItem Value="11" Text="依站別未出險清冊" />
                        <asp:ListItem Value="12" Text="依站別有出險清冊" />
                        <asp:ListItem Value="21" Text="依駕駛員未出險清冊" />
                        <asp:ListItem Value="22" Text="依駕駛員有出險清冊" />
                        <asp:ListItem Value="99" Text="統計報表" Selected="True" />
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuildDate_Search" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eBuildDate_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eBuildDate_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Start_Search" runat="server" CssClass="text-Left-Black" Width="10%" AutoPostBack="True"
                        OnTextChanged="eDepNo_Start_Search_TextChanged" />
                    <asp:Label ID="lbDepNo_Start_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbSplit2" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_End_Search" runat="server" CssClass="text-Left-Black" Width="10%" AutoPostBack="True"
                        OnTextChanged="eDepNo_End_Search_TextChanged" />
                    <asp:Label ID="lbDepNo_End_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDriver_Search" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eDriver_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eDriverName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbPreview" runat="server" CssClass="button-Black" Text="預覽資料" OnClick="bbPreview_Click" Width="90%" />
                </td>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Red" Text="列表" OnClick="bbPrint_Click" Width="90%" Visible="false" />
                </td>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Red" Text="匯出 Excel" OnClick="bbExcel_Click" Width="90%" Visible="false" />
                </td>
                <td class="ColWidth-8Col">
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
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel" GroupingText="資料預覽">
        <asp:GridView ID="gridAnecdoteCaseP_1" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataSourceID="sdsAnecdoteCaseP_1" GridLines="None" PageSize="5">
            <Columns>
                <asp:BoundField DataField="InsuranceState" HeaderText="出險" ReadOnly="True" SortExpression="InsuranceState" />
                <asp:BoundField DataField="DepNo" HeaderText="站別編號" SortExpression="DepNo" />
                <asp:BoundField DataField="DepName" HeaderText="站別" SortExpression="DepName" />
                <asp:BoundField DataField="BuildDate" DataFormatString="{0:d}" HeaderText="日期" SortExpression="BuildDate" />
                <asp:BoundField DataField="Driver" HeaderText="Driver" SortExpression="Driver" Visible="False" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員" SortExpression="DriverName" />
                <asp:BoundField DataField="RelCar_ID" HeaderText="對方車號" SortExpression="RelCar_ID" />
                <asp:BoundField DataField="Relationship" HeaderText="對方姓名" SortExpression="Relationship" />
                <asp:BoundField DataField="CaseOccurrence" HeaderText="肇事經過" SortExpression="CaseOccurrence" />
            </Columns>
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
        <asp:GridView ID="gridAnecdoteCaseP_2" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataSourceID="sdsAnecdoteCaseP_2" GridLines="None" PageSize="5">
            <Columns>
                <asp:BoundField DataField="InsuranceState" HeaderText="出險" ReadOnly="True" SortExpression="InsuranceState" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員" SortExpression="DriverName" />
                <asp:BoundField DataField="Driver" HeaderText="員工編號" SortExpression="Driver" />
                <asp:BoundField DataField="BuildDate" DataFormatString="{0:d}" HeaderText="發生日期" SortExpression="BuildDate" />
                <asp:BoundField DataField="RelCar_ID" HeaderText="對方車號" SortExpression="RelCar_ID" />
                <asp:BoundField DataField="Relationship" HeaderText="對方姓名" SortExpression="Relationship" />
                <asp:BoundField DataField="CaseOccurrence" HeaderText="肇事經過" SortExpression="CaseOccurrence" />
            </Columns>
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
        <asp:GridView ID="gridAnecdoteCaseP_99" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataSourceID="sdsAnecdoteCaseP_99" GridLines="None" PageSize="5">
            <Columns>
                <asp:BoundField DataField="DepCount" HeaderText="站別肇事數" ReadOnly="True" SortExpression="DepCount" />
                <asp:BoundField DataField="InsuCount" HeaderText="部別出險數" ReadOnly="True" SortExpression="InsuCount" />
                <asp:CheckBoxField DataField="HasInsurance" HeaderText="出險" SortExpression="HasInsurance" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="站別" SortExpression="DepName" />
                <asp:BoundField DataField="BuildDate" DataFormatString="{0:d}" HeaderText="發生日期" SortExpression="BuildDate" />
                <asp:BoundField DataField="Driver" HeaderText="Driver" SortExpression="Driver" Visible="False" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員" SortExpression="DriverName" />
                <asp:BoundField DataField="ThirdInsurance" HeaderText="第三責任險" ReadOnly="True" SortExpression="ThirdInsurance" />
                <asp:BoundField DataField="CompInsurance" HeaderText="強制險" ReadOnly="True" SortExpression="CompInsurance" />
                <asp:BoundField DataField="PassengerInsu" HeaderText="乘客險" ReadOnly="True" SortExpression="PassengerInsu" />
            </Columns>
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
    <asp:SqlDataSource ID="sdsAnecdoteCaseP_1" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CASE WHEN a.HasInsurance = 1 THEN '已出險' ELSE '未出險' END AS InsuranceState, a.DepNo, a.DepName, a.BuildDate, a.Driver, a.DriverName, b.RelCar_ID, b.Relationship, a.CaseOccurrence FROM AnecdoteCase AS a LEFT OUTER JOIN AnecdoteCaseB AS b ON b.CaseNo = a.CaseNo WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsAnecdoteCaseP_2" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CASE WHEN a.HasInsurance = 1 THEN '已出險' ELSE '未出險' END AS InsuranceState, a.DriverName, a.BuildDate, a.Driver, b.RelCar_ID, b.Relationship, a.CaseOccurrence FROM AnecdoteCase AS a LEFT OUTER JOIN AnecdoteCaseB AS b ON b.CaseNo = a.CaseNo WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsAnecdoteCaseP_99" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT (SELECT COUNT(CaseNo) AS Expr1 FROM AnecdoteCase WHERE (DepNo = a.DepNo)) AS DepCount, (SELECT COUNT(CaseNo) AS Expr1 FROM AnecdoteCase AS AnecdoteCase_1 WHERE (DepNo = a.DepNo) AND (HasInsurance = a.HasInsurance)) AS InsuCount, a.HasInsurance, a.DepNo, a.DepName, a.BuildDate, a.Driver, a.DriverName, SUM(ISNULL(b.ThirdInsurance, 0)) AS ThirdInsurance, SUM(ISNULL(b.CompInsurance, 0)) AS CompInsurance, SUM(ISNULL(b.PassengerInsu, 0)) AS PassengerInsu FROM AnecdoteCase AS a LEFT OUTER JOIN AnecdoteCaseB AS b ON b.CaseNo = a.CaseNo WHERE (1 &lt;&gt; 1) GROUP BY a.CaseNo, a.DepNo, a.DepName, a.BuildDate, a.Driver, a.DriverName, a.HasInsurance"></asp:SqlDataSource>

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
            <LocalReport ReportPath="Report\AnecdoteCaseP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
