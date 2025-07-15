<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="SearchConsData.aspx.cs" Inherits="TyBus_Intranet_Test_V3.SearchConsData" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link rel="stylesheet" type="text/css" href="Style/MainStyle.css" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>查詢總務耗材料號</title>
</head>
<body>
    <form id="SearchConsDataForm" runat="server">
        <div>
            <asp:Label ID="lbConsType" runat="server" CssClass="text-Right-Blue" Text="類別" Width="80px" />
            <asp:DropDownList ID="ddlConsType" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlConsType_SelectedIndexChanged" Width="60px" />
            <asp:Label ID="lbConsName" runat="server" CssClass="text-Right-Blue" Text="品名" Width="80px" />
            <asp:TextBox ID="eConsName" runat="server" CssClass="text-Left-Black" Width="200px" />
            <asp:Button ID="bbSearch" runat="server" CssClass="button-Blue" Text="查詢" OnClick="bbSearch_Click" Width="120px" />
            <br />
            <asp:Label ID="eConsType" runat="server" Visible="false" />
            <br />
            <asp:GridView ID="gridConsList" runat="server" Width="700px" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="ConsNo" ForeColor="#333333" GridLines="None" OnPageIndexChanging="gridConsList_PageIndexChanging" PageSize="20">
                <AlternatingRowStyle BackColor="White" />
                <Columns>
                    <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                    <asp:BoundField DataField="ConsNo" HeaderText="料號" ReadOnly="True" SortExpression="ConsNo" />
                    <asp:BoundField DataField="ConsName" HeaderText="品名" SortExpression="ConsName" />
                    <asp:BoundField DataField="StockQty" HeaderText="庫存量" SortExpression="StockQty" />
                    <asp:BoundField DataField="ConsSpec" HeaderText="規格" SortExpression="ConsSpec" />
                    <asp:BoundField DataField="AVGPrice" HeaderText="平均進價" SortExpression="AVGPrice" />
                    <asp:BoundField DataField="ConsUnit" HeaderText="單位" SortExpression="ConsUnit" Visible="false" />
                </Columns>
                <EditRowStyle BackColor="#2461BF" />
                <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                <RowStyle BackColor="#EFF3FB" />
                <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                <SortedAscendingCellStyle BackColor="#F5F7FB" />
                <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                <SortedDescendingCellStyle BackColor="#E9EBEF" />
                <SortedDescendingHeaderStyle BackColor="#4870BE" />
            </asp:GridView>
            <asp:Button ID="bbCancel" runat="server" CssClass="button-Red" OnClick="bbCancel_Click" Text="取消" Width="120px" />
            <asp:Button ID="bbOK" runat="server" CssClass="button-Blue" OnClick="bbOK_Click" Text="確定" Width="120px" />
        </div>
    </form>
</body>
</html>
