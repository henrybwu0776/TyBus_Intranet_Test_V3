<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="StockMonth.aspx.cs" Inherits="TyBus_Intranet_Test_V3.StockMonth" %>

<asp:Content ID="StockMonthForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">材料月結存清冊</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalYM_Search" runat="server" CssClass="text-Right-Blue" Text="結存年月：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCalYear_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbCalYear_Search" runat="server" CssClass="text-Left-Black" Text=" 年" Width="10%" />
                    <asp:TextBox ID="eCalMonth_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbCalMonth_Search" runat="server" CssClass="text-Left-Black" Text=" 月" Width="10%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="5">
                    <asp:RadioButtonList ID="rbStockType_Search" runat="server" Width="97%" RepeatColumns="7">
                        <asp:ListItem Value="A" Text="全部" Selected="True" />
                        <asp:ListItem Value="M" Text="材料" />
                        <asp:ListItem Value="O" Text="油料" />
                        <asp:ListItem Value="T" Text="輪胎" />
                        <asp:ListItem Value="MO" Text="材料 + 油料" />
                        <asp:ListItem Value="MT" Text="材料 + 輪胎" />
                        <asp:ListItem Value="OT" Text="油料 + 輪胎" />
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出EXCEL" OnClick="bbExcel_Click" Width="95%" />
                </td>
                <td class="ColWidth-8Col">
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
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridMStockList" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataSourceID="sdsMStockList" GridLines="None" Width="100%">
            <Columns>
                <asp:BoundField DataField="stockno" HeaderText="料號" SortExpression="stockno" />
                <asp:BoundField DataField="name" HeaderText="品名" SortExpression="name" />
                <asp:BoundField DataField="sclass" HeaderText="sclass" SortExpression="sclass" Visible="False" />
                <asp:BoundField DataField="preqty" HeaderText="期初數量" SortExpression="preqty" />
                <asp:BoundField DataField="preupc" HeaderText="期初金額" SortExpression="preupc" />
                <asp:BoundField DataField="qty" HeaderText="期末數量" SortExpression="qty" />
                <asp:BoundField DataField="upc" HeaderText="期末金額" SortExpression="upc" />
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
    <asp:SqlDataSource ID="sdsMStockList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT a.stockno, b.name, b.sclass, a.preqty, a.preupc, a.qty, a.upc FROM MSTOCK AS a LEFT OUTER JOIN STOCK AS b ON b.no = a.stockno WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
