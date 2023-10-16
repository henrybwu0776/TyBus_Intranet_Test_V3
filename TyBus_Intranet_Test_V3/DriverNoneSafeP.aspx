<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DriverNoneSafeP.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DriverNoneSafeP" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="DriverNoneSafePForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">駕駛員未領精勤獎金一覽表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalYM_Search" runat="server" CssClass="text-Right-Blue" Text="查詢年月：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:TextBox ID="eCalYear_Search" runat="server" CssClass="text-Left-Black" Width="20%" />
                    <asp:Label ID="lbCalYear_Search" runat="server" CssClass="text-Left-Black" Text=" 年 " Width="5%" />
                    <asp:TextBox ID="eCalMonth_Start_Search" runat="server" CssClass="text-Left-Black" Width="15%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text=" ～ " Width="5%" />
                    <asp:TextBox ID="eCalMonth_End_Search" runat="server" CssClass="text-Left-Black" Width="15%" />
                    <asp:Label ID="lbCalMonth_Search" runat="server" CssClass="text-Left-Black" Text=" 月 " Width="5%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="單位別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:TextBox ID="eDepNo_Start_Search" runat="server" CssClass="text-Left-Black" OnTextChanged="eDepNo_Start_Search_TextChanged" Width="45%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text=" ～ " Width="5%" />
                    <asp:TextBox ID="eDepNo_End_Search" runat="server" CssClass="text-Left-Black" OnTextChanged="eDepNo_End_Search_TextChanged" Width="45%" />
                    <br />
                    <asp:Label ID="eDepName_Start_Search" runat="server" CssClass="text-Left-Black" Width="50%" />
                    <asp:Label ID="eDepName_End_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="資料查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Black" Text="預覽報表" OnClick="bbPrint_Click" Width="90%" />
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
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridDriverNoneSafeList" runat="server" AllowPaging="True" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataSourceID="sdsDriverNoneSafeList" GridLines="None" AutoGenerateColumns="False" DataKeyNames="EmpNo" PageSize="15" Width="90%">
            <Columns>
                <asp:BoundField DataField="EmpNo" HeaderText="代號" ReadOnly="True" SortExpression="EmpNo" />
                <asp:BoundField DataField="Name" HeaderText="姓名" SortExpression="Name" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="DepGroup" HeaderText="DepGroup" ReadOnly="True" SortExpression="DepGroup" Visible="False" />
                <asp:BoundField DataField="CalMonth" HeaderText="月份" ReadOnly="True" SortExpression="CalMonth" />
                <asp:BoundField DataField="WorkDays" HeaderText="上班天數" ReadOnly="True" SortExpression="WorkDays" />
                <asp:BoundField DataField="TotalKMs" HeaderText="公里數" ReadOnly="True" SortExpression="TotalKMs" />
                <asp:BoundField DataField="SplDays" HeaderText="應休特休" ReadOnly="True" SortExpression="SplDays" />
                <asp:BoundField DataField="UseDays" HeaderText="已休特休" ReadOnly="True" SortExpression="UseDays" />
                <asp:BoundField DataField="ESCDay01" HeaderText="特休" ReadOnly="True" SortExpression="ESCDay01" />
                <asp:BoundField DataField="ESCDay02" HeaderText="公假" ReadOnly="True" SortExpression="ESCDay02" />
                <asp:BoundField DataField="ESCDay03" HeaderText="公傷" ReadOnly="True" SortExpression="ESCDay03" />
                <asp:BoundField DataField="ESCDay04" HeaderText="事假" ReadOnly="True" SortExpression="ESCDay04" />
                <asp:BoundField DataField="ESCDay05" HeaderText="病假" ReadOnly="True" SortExpression="ESCDay05" />
                <asp:BoundField DataField="OtherESCDay" HeaderText="其他" ReadOnly="True" SortExpression="OtherESCDay" />
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
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
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
            <LocalReport ReportPath="Report\DriverNoneSafeP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsDriverNoneSafeList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT EmpNo, Name, DepNo, DepName, DepGroup, CalMonth, WorkDays, TotalKMs, SplDays, UseDays, ESC01 / 8 AS ESCDay01, ESC02 / 8 AS ESCDay02, ESC03 / 8 AS ESCDay03, ESC04 / 8 AS ESCDay04, ESC05 / 8 AS ESCDay05, OtherESC / 8 AS OtherESCDay FROM (SELECT EMPNO, NAME, DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DEPNO)) AS DepName, CAST(1 AS int) AS CalMonth, (SELECT COUNT(ASSIGNNO) AS Expr1 FROM RUNSHEETA WHERE (DRIVER = e.EMPNO) AND (BUDATE BETWEEN '2019/01/01' AND '2019/01/31')) AS WorkDays, ISNULL((SELECT SUM(ActualKm) AS Expr1 FROM RUNSHEETA AS RUNSHEETA_1 WHERE (DRIVER = e.EMPNO) AND (BUDATE BETWEEN '2019/01/01' AND '2019/01/31')), 0) AS TotalKMs, (SELECT depGroup FROM DEPARTMENT AS DEPARTMENT_1 WHERE (DEPNO = e.DEPNO)) AS DepGroup, ISNULL((SELECT spldays FROM YearHoliday WHERE (Years = 2019) AND (EMPNO = e.EMPNO)), 0) AS SplDays, ISNULL((SELECT usedays FROM YearHoliday AS YearHoliday_1 WHERE (Years = 2019) AND (EMPNO = e.EMPNO)), 0) AS UseDays, ISNULL((SELECT SUM(hours) AS Expr1 FROM ESCDUTY WHERE (applyman = e.EMPNO) AND (esctype = '01') AND (realday BETWEEN '2019/01/01' AND '2019/01/31')), 0) AS ESC01, ISNULL((SELECT SUM(hours) AS Expr1 FROM ESCDUTY AS ESCDUTY_5 WHERE (applyman = e.EMPNO) AND (esctype = '02') AND (realday BETWEEN '2019/01/01' AND '2019/01/31')), 0) AS ESC02, ISNULL((SELECT SUM(hours) AS Expr1 FROM ESCDUTY AS ESCDUTY_4 WHERE (applyman = e.EMPNO) AND (esctype = '03') AND (realday BETWEEN '2019/01/01' AND '2019/01/31')), 0) AS ESC03, ISNULL((SELECT SUM(hours) AS Expr1 FROM ESCDUTY AS ESCDUTY_3 WHERE (applyman = e.EMPNO) AND (esctype = '04') AND (realday BETWEEN '2019/01/01' AND '2019/01/31')), 0) AS ESC04, ISNULL((SELECT SUM(hours) AS Expr1 FROM ESCDUTY AS ESCDUTY_2 WHERE (applyman = e.EMPNO) AND (esctype = '05') AND (realday BETWEEN '2019/01/01' AND '2019/01/31')), 0) AS ESC05, ISNULL((SELECT SUM(hours) AS Expr1 FROM ESCDUTY AS ESCDUTY_1 WHERE (applyman = e.EMPNO) AND (esctype NOT IN ('01', '02', '03', '04', '05')) AND (realday BETWEEN '2019/01/01' AND '2019/01/31')), 0) AS OtherESC FROM EMPLOYEE AS e WHERE (TYPE = '20') AND (LEAVEDAY IS NULL OR LEAVEDAY &gt; '2019/01/31')) AS z WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
