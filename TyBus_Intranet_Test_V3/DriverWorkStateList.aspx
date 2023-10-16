<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DriverWorkStateList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DriverWorkStateList" %>
<asp:Content ID="DriverWorkStateListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">駕駛員排班狀況表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbSheetYM" runat="server" CssClass="text-Right-Blue" Text="查詢年月：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eSheetYear" runat="server" CssClass="text-Left-Black" Width="25%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:TextBox ID="eSheetMonth" runat="server" CssClass="text-Left-Black" Width="15%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text=" 月" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_S" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_3" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_E" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
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
        <asp:GridView ID="gridDriverWorkStateList" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataSourceID="sdsDriverWorkStateList" GridLines="None" PageSize="20" Width="100%" OnRowDataBound="gridDriverWorkStateList_RowDataBound">
            <Columns>
                <asp:BoundField DataField="DEPNO" HeaderText="站別編號" SortExpression="DEPNO" />
                <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="DRIVER" HeaderText="駕駛工號" SortExpression="DRIVER" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛姓名" ReadOnly="True" SortExpression="DriverName" />
                <asp:BoundField DataField="WS01" HeaderText="1" ReadOnly="True" SortExpression="WS01" />
                <asp:BoundField DataField="WS02" HeaderText="2" ReadOnly="True" SortExpression="WS02" />
                <asp:BoundField DataField="WS03" HeaderText="3" ReadOnly="True" SortExpression="WS03" />
                <asp:BoundField DataField="WS04" HeaderText="4" ReadOnly="True" SortExpression="WS04" />
                <asp:BoundField DataField="WS05" HeaderText="5" ReadOnly="True" SortExpression="WS05" />
                <asp:BoundField DataField="WS06" HeaderText="6" ReadOnly="True" SortExpression="WS06" />
                <asp:BoundField DataField="WS07" HeaderText="7" ReadOnly="True" SortExpression="WS07" />
                <asp:BoundField DataField="WS08" HeaderText="8" ReadOnly="True" SortExpression="WS08" />
                <asp:BoundField DataField="WS09" HeaderText="9" ReadOnly="True" SortExpression="WS09" />
                <asp:BoundField DataField="WS10" HeaderText="10" ReadOnly="True" SortExpression="WS10" />
                <asp:BoundField DataField="WS11" HeaderText="11" ReadOnly="True" SortExpression="WS11" />
                <asp:BoundField DataField="WS12" HeaderText="12" ReadOnly="True" SortExpression="WS12" />
                <asp:BoundField DataField="WS13" HeaderText="13" ReadOnly="True" SortExpression="WS13" />
                <asp:BoundField DataField="WS14" HeaderText="14" ReadOnly="True" SortExpression="WS14" />
                <asp:BoundField DataField="WS15" HeaderText="15" ReadOnly="True" SortExpression="WS15" />
                <asp:BoundField DataField="WS16" HeaderText="16" ReadOnly="True" SortExpression="WS16" />
                <asp:BoundField DataField="WS17" HeaderText="17" ReadOnly="True" SortExpression="WS17" />
                <asp:BoundField DataField="WS18" HeaderText="18" ReadOnly="True" SortExpression="WS18" />
                <asp:BoundField DataField="WS19" HeaderText="19" ReadOnly="True" SortExpression="WS19" />
                <asp:BoundField DataField="WS20" HeaderText="20" ReadOnly="True" SortExpression="WS20" />
                <asp:BoundField DataField="WS21" HeaderText="21" ReadOnly="True" SortExpression="WS21" />
                <asp:BoundField DataField="WS22" HeaderText="22" ReadOnly="True" SortExpression="WS22" />
                <asp:BoundField DataField="WS23" HeaderText="23" ReadOnly="True" SortExpression="WS23" />
                <asp:BoundField DataField="WS24" HeaderText="24" ReadOnly="True" SortExpression="WS24" />
                <asp:BoundField DataField="WS25" HeaderText="25" ReadOnly="True" SortExpression="WS25" />
                <asp:BoundField DataField="WS26" HeaderText="26" ReadOnly="True" SortExpression="WS26" />
                <asp:BoundField DataField="WS27" HeaderText="27" ReadOnly="True" SortExpression="WS27" />
                <asp:BoundField DataField="WS28" HeaderText="28" ReadOnly="True" SortExpression="WS28" />
                <asp:BoundField DataField="WS29" HeaderText="29" ReadOnly="True" SortExpression="WS29" />
                <asp:BoundField DataField="WS30" HeaderText="30" ReadOnly="True" SortExpression="WS30" />
                <asp:BoundField DataField="WS31" HeaderText="31" ReadOnly="True" SortExpression="WS31" />
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
    <asp:SqlDataSource ID="sdsDriverWorkStateList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT DEPNO, CAST('' AS varchar) AS DepName, DRIVER, CAST('' AS varchar) AS DriverName, CAST('' AS varchar) AS WS01, CAST('' AS varchar) AS WS02, CAST('' AS varchar) AS WS03, CAST('' AS varchar) AS WS04, CAST('' AS varchar) AS WS05, CAST('' AS varchar) AS WS06, CAST('' AS varchar) AS WS07, CAST('' AS varchar) AS WS08, CAST('' AS varchar) AS WS09, CAST('' AS varchar) AS WS10, CAST('' AS varchar) AS WS11, CAST('' AS varchar) AS WS12, CAST('' AS varchar) AS WS13, CAST('' AS varchar) AS WS14, CAST('' AS varchar) AS WS15, CAST('' AS varchar) AS WS16, CAST('' AS varchar) AS WS17, CAST('' AS varchar) AS WS18, CAST('' AS varchar) AS WS19, CAST('' AS varchar) AS WS20, CAST('' AS varchar) AS WS21, CAST('' AS varchar) AS WS22, CAST('' AS varchar) AS WS23, CAST('' AS varchar) AS WS24, CAST('' AS varchar) AS WS25, CAST('' AS varchar) AS WS26, CAST('' AS varchar) AS WS27, CAST('' AS varchar) AS WS28, CAST('' AS varchar) AS WS29, CAST('' AS varchar) AS WS30, CAST('' AS varchar) AS WS31 FROM RUNSHEETA WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
