<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="SearchProduct.aspx.cs" Inherits="TyBus_Intranet_Test_V3.SearchProduct" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link rel="stylesheet" type="text/css" href="Style/MainStyle.css" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>查詢耗材料號</title>
</head>
<body>
    <form id="SearchProductForm" runat="server">
        <div>
            <asp:Label ID="lbProductName_Search" runat="server" CssClass="text-Right-Blue" Text="品名" Width="120px" />
            <asp:TextBox ID="eProductName_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eProductName_Search_TextChanged" Width="400px" />
            <asp:GridView ID="gridProductData_Search" runat="server" CssClass="text-Left-Black" Width="600px" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="ProductNo" DataSourceID="dsProduct" ForeColor="#333333" GridLines="None">
                <AlternatingRowStyle BackColor="White" />
                <Columns>
                    <asp:CommandField ShowSelectButton="True" ButtonType="Button" />
                    <asp:BoundField DataField="ProductNo" HeaderText="料號" ReadOnly="True" SortExpression="ProductNo" />
                    <asp:BoundField DataField="ProductName" HeaderText="品名" SortExpression="ProductName" ReadOnly="True" />
                    <asp:BoundField DataField="Price" HeaderText="單價" />
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
            <asp:Button ID="bbCancel" runat="server" CssClass="button-Red" OnClick="bbCancel_Click" Text="取消" Width="120px" />
            <asp:Button ID="bbOK" runat="server" CssClass="button-Blue" OnClick="bbOK_Click" Text="確定" Width="120px" />
        </div>
    </form>
    <asp:SqlDataSource ID="dsProduct" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT ConsNo AS ProductNo, isnull(CorrespondModel, '') + '-' + isnull(ConsName, '') + '(' + isnull(Spec_Color, '') + ')' AS ProductName, Price FROM IAConsumables WHERE (ISNULL(ConsNo, '') &lt;&gt; '')"></asp:SqlDataSource>
</body>
</html>
