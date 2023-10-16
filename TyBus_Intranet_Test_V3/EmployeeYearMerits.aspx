<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmployeeYearMerits.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmployeeYearMerits" %>
<asp:Content ID="EmployeeYearMeritsForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">員工考核參考資料查詢</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCheckYear" runat="server" CssClass="text-Right-Blue" Text="查詢年度：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCheckYear" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="95%" />
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
    <asp:GridView ID="gridDataList" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" CellPadding="3" DataSourceID="sdsDataList" GridLines="Horizontal" Width="100%">
        <AlternatingRowStyle BackColor="#F7F7F7" />
        <Columns>
            <asp:BoundField DataField="DEPNO" HeaderText="DEPNO" ReadOnly="True" SortExpression="DEPNO" Visible="False" />
            <asp:BoundField DataField="DepName" HeaderText="單位" ReadOnly="True" SortExpression="DepName" />
            <asp:BoundField DataField="TITLE" HeaderText="TITLE" ReadOnly="True" SortExpression="TITLE" Visible="False" />
            <asp:BoundField DataField="Title_C" HeaderText="職稱" ReadOnly="True" SortExpression="Title_C" />
            <asp:BoundField DataField="ASSUMEDAY" DataFormatString="{0:D}" HeaderText="到職日" ReadOnly="True" SortExpression="ASSUMEDAY" />
            <asp:BoundField DataField="EMPNO" HeaderText="代號" ReadOnly="True" SortExpression="EMPNO" />
            <asp:BoundField DataField="NAME" HeaderText="姓名" ReadOnly="True" SortExpression="NAME" />
            <asp:BoundField DataField="salarypoint" HeaderText="俸點" ReadOnly="True" SortExpression="salarypoint" />
            <asp:BoundField DataField="TruNum" HeaderText="曠職" ReadOnly="True" SortExpression="TruNum" />
            <asp:BoundField DataField="EscType4" HeaderText="事假" ReadOnly="True" SortExpression="EscType4" />
            <asp:BoundField DataField="EscType5" HeaderText="病假" ReadOnly="True" SortExpression="EscType5" />
            <asp:BoundField DataField="EscType2" HeaderText="公假" ReadOnly="True" SortExpression="EscType2" />
            <asp:BoundField DataField="Warning" HeaderText="警告" ReadOnly="True" SortExpression="Warning" />
            <asp:BoundField DataField="Demerit3" HeaderText="申誡" ReadOnly="True" SortExpression="Demerit3" />
            <asp:BoundField DataField="Demerit2" HeaderText="記過" ReadOnly="True" SortExpression="Demerit2" />
            <asp:BoundField DataField="Demerit1" HeaderText="記大過" ReadOnly="True" SortExpression="Demerit1" />
            <asp:BoundField DataField="Merit3" HeaderText="嘉獎" ReadOnly="True" SortExpression="Merit3" />
            <asp:BoundField DataField="Merit2" HeaderText="小功" ReadOnly="True" SortExpression="Merit2" />
            <asp:BoundField DataField="Merit1" HeaderText="記大功" ReadOnly="True" SortExpression="Merit1" />
        </Columns>
        <FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
        <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#F7F7F7" />
        <PagerStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" HorizontalAlign="Right" />
        <RowStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" />
        <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="#F7F7F7" />
        <SortedAscendingCellStyle BackColor="#F4F4FD" />
        <SortedAscendingHeaderStyle BackColor="#5A4C9D" />
        <SortedDescendingCellStyle BackColor="#D8D8F0" />
        <SortedDescendingHeaderStyle BackColor="#3E3277" />
    </asp:GridView>
    <asp:SqlDataSource ID="sdsDataList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = z.DEPNO)) AS DepName, TITLE, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '員工資料        EMPLOYEE        TITLE') AND (CLASSNO = z.TITLE)) AS Title_C, ASSUMEDAY, EMPNO, NAME, salarypoint, SUM(TruNum) AS TruNum, SUM(EscType4) AS EscType4, SUM(EscType5) AS EscType5, SUM(EscType2) AS EscType2, SUM(Warning) AS Warning, SUM(Demerit3) AS Demerit3, SUM(Demerit2) AS Demerit2, SUM(Demerit1) AS Demerit1, SUM(Merit3) AS Merit3, SUM(Merit2) AS Merit2, SUM(Merit1) AS Merit1 FROM (SELECT a.DEPNO, a.TITLE, a.ASSUMEDAY, a.EMPNO, a.NAME, a.salarypoint, ISNULL(b.trunum, 0) AS TruNum, ISNULL(b.esctype4, 0) AS EscType4, ISNULL(b.esctype5, 0) AS EscType5, ISNULL(b.esctype2, 0) AS EscType2, CAST(0 AS float) AS Warning, CAST(0 AS float) AS Demerit3, CAST(0 AS float) AS Demerit2, CAST(0 AS float) AS Demerit1, CAST(0 AS float) AS Merit3, CAST(0 AS float) AS Merit2, CAST(0 AS float) AS Merit1 FROM EMPLOYEE AS a LEFT OUTER JOIN MHGS AS b ON b.empno = a.EMPNO WHERE (a.DEPNO &lt;&gt; '00') AND (a.LEAVEDAY IS NULL) AND (YEAR(b.attmonth) = @Year) UNION ALL SELECT a.DEPNO, a.TITLE, a.ASSUMEDAY, a.EMPNO, a.NAME, a.salarypoint, CAST(0 AS float) AS TruNum, CAST(0 AS float) AS EscType4, CAST(0 AS float) AS EscType5, CAST(0 AS float) AS EscType2, ISNULL(c.WARNING, 0) AS Warning, ISNULL(c.DEMERIT3, 0) AS Demerit3, ISNULL(c.DEMERIT2, 0) AS Demerit2, ISNULL(c.DEMERIT1, 0) AS Demerit1, ISNULL(c.MERIT3, 0) AS Merit3, ISNULL(c.MERIT2, 0) AS Merit2, ISNULL(c.MERIT1, 0) AS Merit1 FROM EMPLOYEE AS a LEFT OUTER JOIN MERITSA AS c ON c.EMPNO = a.EMPNO WHERE (a.DEPNO &lt;&gt; '00') AND (a.LEAVEDAY IS NULL) AND (YEAR(c.MERITYEAR) = @Year)) AS z GROUP BY DEPNO, TITLE, ASSUMEDAY, EMPNO, NAME, salarypoint ORDER BY DEPNO, TITLE">
        <SelectParameters>
            <asp:Parameter Name="Year" />
        </SelectParameters>
    </asp:SqlDataSource>
</asp:Content>
