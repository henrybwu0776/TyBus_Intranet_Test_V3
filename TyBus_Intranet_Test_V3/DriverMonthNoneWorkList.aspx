<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DriverMonthNoneWorkList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DriverMonthNoneWorkList" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="DriverMonthNoneWorkListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">全月未出勤駕駛員查詢</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalYM" runat="server" CssClass="text-Right-Blue" Text="查詢年月：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCalYear" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:TextBox ID="eCalMonth" runat="server" CssClass="text-Left-Black" Width="20%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text=" 月" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_TextChanged" Width="30%" />
                    <asp:Label ID="eDepName" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbEmpNo" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eEmpNo" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_TextChanged" Width="30%" />
                    <asp:Label ID="eEmpName" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Blue" Text="列印" OnClick="bbPrint_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出EXCEL" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="離開" OnClick="bbClose_Click" Width="90%" />
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
        <asp:GridView ID="gridShowData" runat="server" AutoGenerateColumns="False" CellPadding="4" DataSourceID="sdsShowData" ForeColor="#333333" GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:BoundField DataField="DEPNO" HeaderText="DEPNO" ReadOnly="True" SortExpression="DEPNO" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="單位" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="NAME" HeaderText="姓名" ReadOnly="True" SortExpression="NAME" />
                <asp:BoundField DataField="EMPNO" HeaderText="工號" ReadOnly="True" SortExpression="EMPNO" />
                <asp:BoundField DataField="worktype" HeaderText="在職狀況" ReadOnly="True" SortExpression="worktype" />
                <asp:BoundField DataField="LEAVEDAY" DataFormatString="{0:D}" HeaderText="離職日" ReadOnly="True" SortExpression="LEAVEDAY" />
                <asp:BoundField DataField="BEGINDAY" DataFormatString="{0:D}" HeaderText="留停起" ReadOnly="True" SortExpression="BEGINDAY" />
                <asp:BoundField DataField="STOPDAY" DataFormatString="{0:D}" HeaderText="留停迄" ReadOnly="True" SortExpression="STOPDAY" />
                <asp:BoundField DataField="TDay01" HeaderText="特休" ReadOnly="True" SortExpression="TDay01" />
                <asp:BoundField DataField="TDay03" HeaderText="公傷假" ReadOnly="True" SortExpression="TDay03" />
                <asp:BoundField DataField="TDay04" HeaderText="事假" ReadOnly="True" SortExpression="TDay04" />
                <asp:BoundField DataField="TDay05" HeaderText="病假" ReadOnly="True" SortExpression="TDay05" />
                <asp:BoundField DataField="TOtherDays" HeaderText="其他假別" ReadOnly="True" SortExpression="TOtherDays" />
                <asp:BoundField DataField="TotalESCDays" HeaderText="總請假天數" ReadOnly="True" SortExpression="總請假天數" />
            </Columns>
            <EditRowStyle BackColor="#7C6F57" />
            <FooterStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#666666" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#E3EAEB" />
            <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#F8FAFA" />
            <SortedAscendingHeaderStyle BackColor="#246B61" />
            <SortedDescendingCellStyle BackColor="#D4DFE1" />
            <SortedDescendingHeaderStyle BackColor="#15524A" />
        </asp:GridView>
        <asp:SqlDataSource ID="sdsShowData" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = z.DEPNO)) AS DepName, EMPNO, NAME, worktype, LEAVEDAY, BEGINDAY, STOPDAY, SUM(Hours01) / 8 AS TDay01, SUM(Hours03) / 8 AS TDay03, SUM(Hours04) / 8 AS TDay04, SUM(Hours05) / 8 AS TDay05, SUM(OtherHours) / 8 AS TOtherDays, CASE WHEN SUM(z.Hours01 + z.Hours03 + z.Hours04 + z.Hours05 + z.OtherHours) = 0 THEN '無請假記錄' ELSE CAST(SUM(z.Hours01 + z.Hours03 + z.Hours04 + z.Hours05 + z.OtherHours) / 8 AS varchar) END AS TotalESCDays FROM (SELECT DEPNO, EMPNO, NAME, LEAVEDAY, BEGINDAY, STOPDAY, worktype, CASE WHEN t.ESCType = '01' THEN t.[Hours] ELSE 0 END AS Hours01, CASE WHEN t.ESCType = '03' THEN t.[Hours] ELSE 0 END AS Hours03, CASE WHEN t.ESCType = '04' THEN t.[Hours] ELSE 0 END AS Hours04, CASE WHEN t.ESCType = '05' THEN t.[Hours] ELSE 0 END AS Hours05, CASE WHEN t.ESCType NOT IN ('01' , '03' , '04' , '05') THEN t.[Hours] ELSE 0 END AS OtherHours FROM (SELECT e.DEPNO, e.EMPNO, e.NAME, e.worktype, e.LEAVEDAY, e.BEGINDAY, e.STOPDAY, d.esctype, d.hours FROM EMPLOYEE AS e LEFT OUTER JOIN ESCDUTY AS d ON d.applyman = e.EMPNO AND d.applyday BETWEEN '2020/04/01' AND '2020/04/30' WHERE (e.LEAVEDAY BETWEEN '2020/04/01' AND '2020/04/30') AND (e.TYPE = '20') AND (e.EMPNO NOT IN (SELECT DISTINCT DRIVER FROM RUNSHEETA WHERE (BUDATE BETWEEN '2020/04/01' AND '2020/04/30'))) AND (e.DEPNO IN (SELECT DEPNO FROM DEPARTMENT AS DEPARTMENT_3 WHERE (ISSTATION = 'V') AND (depGroup IN ('1', '2')) AND (PATINDEX('Z%', DEPNO) = 0))) UNION ALL SELECT e.DEPNO, e.EMPNO, e.NAME, e.worktype, e.LEAVEDAY, e.BEGINDAY, e.STOPDAY, d.esctype, d.hours FROM EMPLOYEE AS e LEFT OUTER JOIN ESCDUTY AS d ON d.applyman = e.EMPNO AND d.applyday BETWEEN '2020/04/01' AND '2020/04/30' WHERE (e.BEGINDAY BETWEEN '2020/04/01' AND '2020/04/30') AND (e.STOPDAY IS NULL OR e.STOPDAY &gt; '2020/04/01') AND (e.TYPE = '20') AND (e.EMPNO NOT IN (SELECT DISTINCT DRIVER FROM RUNSHEETA AS RUNSHEETA_2 WHERE (BUDATE BETWEEN '2020/04/01' AND '2020/04/30'))) AND (e.DEPNO IN (SELECT DEPNO FROM DEPARTMENT AS DEPARTMENT_2 WHERE (ISSTATION = 'V') AND (depGroup IN ('1', '2')) AND (PATINDEX('Z%', DEPNO) = 0))) UNION ALL SELECT e.DEPNO, e.EMPNO, e.NAME, e.worktype, e.LEAVEDAY, e.BEGINDAY, e.STOPDAY, d.esctype, d.hours FROM EMPLOYEE AS e LEFT OUTER JOIN ESCDUTY AS d ON d.applyman = e.EMPNO AND d.applyday BETWEEN '2020/04/01' AND '2020/04/30' WHERE (e.BEGINDAY IS NULL) AND (e.LEAVEDAY IS NULL) AND (e.TYPE = '20') AND (e.EMPNO NOT IN (SELECT DISTINCT DRIVER FROM RUNSHEETA AS RUNSHEETA_1 WHERE (BUDATE BETWEEN '2020/04/01' AND '2020/04/30'))) AND (e.DEPNO IN (SELECT DEPNO FROM DEPARTMENT AS DEPARTMENT_1 WHERE (ISSTATION = 'V') AND (depGroup IN ('1', '2')) AND (PATINDEX('Z%', DEPNO) = 0)))) AS t) AS z WHERE (1 &lt;&gt; 1) GROUP BY DEPNO, EMPNO, NAME, LEAVEDAY, BEGINDAY, STOPDAY, worktype ORDER BY EMPNO"></asp:SqlDataSource>
    </asp:Panel>
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCLoseReport" runat="server" CssClass="button-Red" Text="關閉報表" OnClick="bbCLoseReport_Click" Width="120px" />
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
            <LocalReport ReportPath="Report\DriverMonthNoneWorkList.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
