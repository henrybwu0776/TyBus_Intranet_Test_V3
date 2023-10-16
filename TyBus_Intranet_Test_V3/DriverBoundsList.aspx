<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DriverBoundsList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DriverBoundsList" %>

<asp:Content ID="DriverBoundsListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">駕駛員大巴／國道／AB班趟次表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eDepName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalYM_Search" runat="server" CssClass="text-Right-Blue" Text="統計年月：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCalYear_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:TextBox ID="eCalMonth_Search" runat="server" CssClass="text-Left-Black" Width="25%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text=" 月份" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbShowData" runat="server" CssClass="button-Black" OnClick="bbShowData_Click" Text="產生清冊" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" OnClick="bbExcel_Click" Text="匯出 EXCEL" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" OnClick="bbClose_Click" Text="結束" Width="90%" />
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
        <asp:Label ID="lbWarningTest" runat="server" CssClass="errorMessageText" Text="本列表來源為行車憑單，僅供參考；一旦行車憑單有任何變更時，需重新產生清冊" />
        <asp:GridView ID="gridShowData" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataSourceID="sdsShowData" GridLines="None" Width="100%">
            <Columns>
                <asp:BoundField DataField="CalYM" HeaderText="統計年月" ReadOnly="True" SortExpression="CalYM" />
                <asp:BoundField DataField="DEPNO" HeaderText="站別編號" SortExpression="DEPNO" />
                <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="DRIVER" HeaderText="駕駛員工號" SortExpression="DRIVER" />
                <asp:BoundField DataField="DriverName" HeaderText="姓名" ReadOnly="True" SortExpression="DriverName" />
                <asp:BoundField DataField="TotalBounds_R" HeaderText="AB班趟次" ReadOnly="True" SortExpression="TotalBounds_R" />
                <asp:BoundField DataField="TotalBounds_L" HeaderText="國道趟次" ReadOnly="True" SortExpression="TotalBounds_L" />
                <asp:BoundField DataField="TotalBounds_C" HeaderText="大巴趟次" ReadOnly="True" SortExpression="TotalBounds_C" />
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
    <asp:SqlDataSource ID="sdsShowData" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CAST('' AS nvarchar(12)) AS CalYM, DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = t.DEPNO)) AS DepName, DRIVER, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = t.DRIVER)) AS DriverName, SUM(RunItemBounds) AS TotalBounds_R, SUM(LinesBounds) AS TotalBounds_L, SUM(Car_Class) AS TotalBounds_C FROM (SELECT a.DEPNO, a.DRIVER, CASE WHEN isnull((SELECT HasBounds FROM RunItem r WHERE DepNo = a.DepNo AND RunItemNo = b.RunItem) , '') = 'V' THEN 1 ELSE 0 END AS RunItemBounds, CASE WHEN isnull((SELECT HasBounds FROM Lines l WHERE l.LinesNo = b.LinesNo) , '') = 'V' THEN 1 ELSE 0 END AS LinesBounds, CASE WHEN isnull((SELECT Class FROM Car_InfoA WHERE Car_ID = b.Car_ID) , '') = '甲' THEN 1 ELSE 0 END AS Car_Class FROM RunSheetB AS b LEFT OUTER JOIN RUNSHEETA AS a ON a.ASSIGNNO = b.AssignNo WHERE (a.BUDATE BETWEEN '2019/12/01' AND '2019/12/31') AND (ISNULL(b.ReduceReason, '') = '')) AS t WHERE (1 &lt;&gt; 1) GROUP BY DEPNO, DRIVER"></asp:SqlDataSource>
</asp:Content>
