<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DriverWorkState.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DriverWorkState" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="DriverWorkStateForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">駕駛員出勤狀況查核</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="查詢站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_S_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_S_Search_TextChanged" Width="35%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text=" ～ " />
                    <asp:TextBox ID="eDepNo_E_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_E_Search_TextChanged" Width="35%" />
                    <br />
                    <asp:Label ID="eDepName_S_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="eDepName_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalYM_Search" runat="server" CssClass="text-Right-Blue" Text="計算年月：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCalYear_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:TextBox ID="eCalMonth_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="lbSplit_3" runat="server" CssClass="text-Left-Black" Text=" 月" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDriver_Search" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDriver_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Search_TextChanged" Width="45%" />
                    <asp:Label ID="eDriverName_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="異常查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbAll2Excel" runat="server" CssClass="button-Black" Text="全部匯出" OnClick="bbAll2Excel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbError2Excel" runat="server" CssClass="button-Red" Text="異常匯出" OnClick="bbError2Excel_Click" Width="90%" />
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
    <asp:Panel ID="plShowData" runat="server" Width="100%">
        <asp:Label ID="lbRemark" runat="server" CssClass="errorMessageText" Text="備註：新進、退休、離職、停職及未滿一個月之人員，請用人工核定！" Width="97%" />
        <asp:GridView ID="gridDriverWorkStateErrorList" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="EMPNO" DataSourceID="sdsDriverWorkStateErrorList" GridLines="None">
            <Columns>
                <asp:BoundField DataField="DEPNO" HeaderText="DEPNO" SortExpression="DEPNO" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="EMPNO" HeaderText="駕駛員工號" ReadOnly="True" SortExpression="EMPNO" />
                <asp:BoundField DataField="NAME" HeaderText="駕駛員" SortExpression="NAME" />
                <asp:BoundField DataField="AssumeDay" HeaderText="到職日期" SortExpression="worktype" DataFormatString="{0:D}" />
                <asp:BoundField DataField="WorkType" HeaderText="在職情況" SortExpression="worktype" />
                <asp:BoundField DataField="MonthDays_Cal" HeaderText="MonthDays_Cal" ReadOnly="True" SortExpression="MonthDays_Cal" Visible="False" />
                <asp:BoundField DataField="MonthDays" HeaderText="當月天數" ReadOnly="True" SortExpression="MonthDays" />
                <asp:BoundField DataField="WorkDays" HeaderText="憑單天數" ReadOnly="True" SortExpression="WorkDays" />
                <asp:BoundField DataField="ESCDays" HeaderText="請假天數" ReadOnly="True" SortExpression="ESCDays" />
                <asp:BoundField DataField="MonthAllowESCDays" HeaderText="當月可休天數" ReadOnly="True" SortExpression="MonthAllowESCDays" />
                <asp:BoundField DataField="NoneWorkDays" HeaderText="憑單未出勤天數" ReadOnly="True" SortExpression="NoneWorkDays" />
                <asp:BoundField DataField="NormalDays" HeaderText="憑單平常日天數" ReadOnly="True" SortExpression="NormalDays" />
                <asp:BoundField DataField="AllowNormalDays" HeaderText="AllowNormalDays" ReadOnly="True" SortExpression="AllowNormalDays" Visible="False" />
                <asp:BoundField DataField="HoliDays" HeaderText="休假上班天數" ReadOnly="True" SortExpression="HoliDays" />
                <asp:BoundField DataField="AllowHoliDays" HeaderText="當月應休假天數" ReadOnly="True" SortExpression="AllowHoliDays" />
                <asp:BoundField DataField="RuleHoliDays" HeaderText="例假上班天數" ReadOnly="True" SortExpression="RuleHoliDays" />
                <asp:BoundField DataField="AllowRuleHoliDays" HeaderText="當月應休例假天數" ReadOnly="True" SortExpression="AllowRuleHoliDays" />
                <asp:BoundField DataField="NationHoliDays" HeaderText="國定假日上班天數" ReadOnly="True" SortExpression="NationHoliDays" />
                <asp:BoundField DataField="AllowNationHoliDays" HeaderText="當月國定假日天數" ReadOnly="True" SortExpression="AllowNationHoliDays" />
                <asp:BoundField DataField="HoliDayDiff" HeaderText="休假天數差異" ReadOnly="True" SortExpression="HoliDayDiff" />
                <asp:BoundField DataField="RuleHoliDatDiff" HeaderText="例假天數差異" ReadOnly="True" SortExpression="RuleHoliDatDiff" />
                <asp:BoundField DataField="NationHoliDayDiff" HeaderText="國定假日天數差異" ReadOnly="True" SortExpression="NationHoliDayDiff" />
                <asp:BoundField DataField="NoneWorkDayDiff" HeaderText="當月請休假加班天數差異" ReadOnly="True" SortExpression="NoneWorkDayDiff" />
                <asp:BoundField DataField="MonthDayDiff" HeaderText="當月總天數差異" ReadOnly="True" SortExpression="MonthDayDiff" />
            </Columns>
            <EmptyDataTemplate>
                <asp:Label ID="lbEmptyMessage" runat="server" CssClass="errorMessageText" Text="無異常資料"></asp:Label>
            </EmptyDataTemplate>
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
    <asp:SqlDataSource ID="sdsDriverWorkStateErrorList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = t.DEPNO)) AS DepName, EMPNO, NAME, worktype, NormalDays + HoliDays + RuleHoliDays + NationHoliDays + ESCDays AS MonthDays_Cal, MonthDays, WorkDays, ESCDays, AllowHoliDays + AllowRuleHoliDays + AllowNationHoliDays AS MonthAllowESCDays, NoneWorkDays, NormalDays, AllowNormalDays, HoliDays, AllowHoliDays, RuleHoliDays, AllowRuleHoliDays, NationHoliDays, AllowNationHoliDays, CASE WHEN (t.HoliDays - t.AllowHoliDays) &gt; 0 THEN (t.HoliDays - t.AllowHoliDays) ELSE 0 END AS HoliDayDiff, CASE WHEN (t.RuleHoliDays - t.AllowRuleHoliDays) &gt; 0 THEN (t.RuleHoliDays - t.AllowRuleHoliDays) ELSE 0 END AS RuleHoliDatDiff, CASE WHEN (t.NationHoliDays - t.AllowNationHoliDays) &gt; 0 THEN (t.NationHoliDays - t.AllowNationHoliDays) ELSE 0 END AS NationHoliDayDiff, CASE WHEN ((t.HoliDays + t.RuleHoliDays + t.NationHoliDays + t.NoneWorkDays) &lt;&gt; (t.AllowHoliDays + t.AllowRuleHoliDays + t.AllowNationHoliDays + t.ESCDays)) THEN ((t.HoliDays + t.RuleHoliDays + t.NationHoliDays + t.NoneWorkDays) - (t.AllowHoliDays + t.AllowRuleHoliDays + t.AllowNationHoliDays + t.ESCDays)) ELSE 0 END AS NoneWorkDayDiff, CASE WHEN ((t.NoneWorkDays + t.NormalDays + t.HoliDays + t.RuleHoliDays + t.NationHoliDays) &lt;&gt; t.MonthDays) THEN (t.NoneWorkDays + t.NormalDays + t.HoliDays + t.RuleHoliDays + t.NationHoliDays) - t.MonthDays ELSE 0 END AS MonthDayDiff FROM (SELECT DAY(DATEADD(m, 1, '2019/02/01') - DAY('2019/02/01')) AS MonthDays, worktype, EMPNO, NAME, DEPNO, (SELECT COUNT(ASSIGNNO) AS Expr1 FROM RUNSHEETA WHERE (BUDATE BETWEEN '2019/02/01' AND '2019/02/28') AND (DRIVER = e.EMPNO)) AS WorkDays, DAY(DATEADD(m, 1, '2019/02/01') - DAY('2019/02/01')) - (SELECT COUNT(ASSIGNNO) AS Expr1 FROM RUNSHEETA AS RUNSHEETA_6 WHERE (BUDATE BETWEEN '2019/02/01' AND '2019/02/28') AND (DRIVER = e.EMPNO)) AS NoneWorkDays, (SELECT COUNT(ASSIGNNO) AS Expr1 FROM RUNSHEETA AS RUNSHEETA_5 WHERE (BUDATE BETWEEN '2019/02/01' AND '2019/02/28') AND (DRIVER = e.EMPNO) AND (WORKSTATE = '0')) AS NormalDays, (SELECT WORKSTATEDAYS0 FROM DriverMonthDays WHERE (YYYYMM = '201902')) AS AllowNormalDays, (SELECT COUNT(ASSIGNNO) AS Expr1 FROM RUNSHEETA AS RUNSHEETA_4 WHERE (BUDATE BETWEEN '2019/02/01' AND '2019/02/28') AND (DRIVER = e.EMPNO) AND (WORKSTATE = '1')) AS HoliDays, (SELECT WORKSTATEDAYS1 FROM DriverMonthDays AS DriverMonthDays_3 WHERE (YYYYMM = '201902')) AS AllowHoliDays, (SELECT COUNT(ASSIGNNO) AS Expr1 FROM RUNSHEETA AS RUNSHEETA_3 WHERE (BUDATE BETWEEN '2019/02/01' AND '2019/02/28') AND (DRIVER = e.EMPNO) AND (WORKSTATE = '2')) AS RuleHoliDays, (SELECT WORKSTATEDAYS2 FROM DriverMonthDays AS DriverMonthDays_2 WHERE (YYYYMM = '201902')) AS AllowRuleHoliDays, (SELECT COUNT(ASSIGNNO) AS Expr1 FROM RUNSHEETA AS RUNSHEETA_2 WHERE (BUDATE BETWEEN '2019/02/01' AND '2019/02/28') AND (DRIVER = e.EMPNO) AND (WORKSTATE = '3')) AS NationHoliDays, (SELECT WORKSTATEDAYS3 FROM DriverMonthDays AS DriverMonthDays_1 WHERE (YYYYMM = '201902')) AS AllowNationHoliDays, (SELECT COUNT(ASSIGNNO) AS Expr1 FROM RUNSHEETA AS RUNSHEETA_1 WHERE (BUDATE BETWEEN '2019/02/01' AND '2019/02/28') AND (DRIVER = e.EMPNO) AND (WORKSTATE &lt;&gt; '0')) AS TotalHoliDays, (SELECT COUNT(Hours_T) AS Expr1 FROM (SELECT SUM(hours) AS Hours_T FROM ESCDUTY WHERE (realday BETWEEN '2019/02/01' AND '2019/02/28') AND (applyman = e.EmpNo) GROUP BY realday) AS z WHERE (Hours_T &gt;= 8)) AS ESCDays FROM EMPLOYEE AS e WHERE (1 &lt;&gt; 1)) AS t"></asp:SqlDataSource>
</asp:Content>
