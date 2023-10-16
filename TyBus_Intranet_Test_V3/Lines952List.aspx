<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="Lines952List.aspx.cs" Inherits="TyBus_Intranet_Test_V3.Lines952List" %>

<asp:Content ID="Lines952ListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">路線匯出新北動態系統用資料</a>
    </div>
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbLinesNo" runat="server" CssClass="text-Right-Blue" Text="路線代號：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eLinesNo" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuDate" runat="server" CssClass="text-Right-Blue" Text="行車日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eBuDate" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出EXCEL" OnClick="bbExcel_Click" Width="90%" />
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
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridDataShow" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataSourceID="sdsDataShow" GridLines="None" Width="90%">
            <Columns>
                <asp:BoundField DataField="DriveDate" DataFormatString="{0:d}" HeaderText="行車日期" ReadOnly="True" SortExpression="DriveDate" />
                <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="DriverNo" HeaderText="駕駛員工號" ReadOnly="True" SortExpression="DriverNo" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員姓名" ReadOnly="True" SortExpression="DriverName" />
                <asp:BoundField DataField="LinesNo" HeaderText="路線代號" ReadOnly="True" SortExpression="LinesNo" />
                <asp:BoundField DataField="Car_Id" HeaderText="牌照號碼" SortExpression="Car_Id" />
                <asp:BoundField DataField="ToTime" HeaderText="去程時刻" SortExpression="ToTime" />
                <asp:BoundField DataField="BackTime" HeaderText="回程時刻" SortExpression="BackTime" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" />
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
    <asp:SqlDataSource ID="sdsDataShow" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT (SELECT BUDATE FROM RUNSHEETA WHERE (ASSIGNNO = b.AssignNo)) AS DriveDate, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = (SELECT DEPNO FROM RUNSHEETA AS RUNSHEETA_3 WHERE (ASSIGNNO = b.AssignNo)))) AS DepName, (SELECT DRIVER FROM RUNSHEETA AS RUNSHEETA_2 WHERE (ASSIGNNO = b.AssignNo)) AS DriverNo, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = (SELECT DRIVER FROM RUNSHEETA AS RUNSHEETA_1 WHERE (ASSIGNNO = b.AssignNo)))) AS DriverName, '952' AS LinesNo, Car_Id, ToTime, BackTime, Remark FROM RunSheetB AS b WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
