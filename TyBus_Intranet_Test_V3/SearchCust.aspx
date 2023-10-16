<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="SearchCust.aspx.cs" Inherits="TyBus_Intranet_Test_V3.SearchCust" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link rel="stylesheet" type="text/css" href="Style/MainStyle.css" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>查詢客戶名稱</title>
</head>
<body>
    <form id="SearchCustForm" runat="server">
        <div>
            <asp:Label ID="lbCustName_Search" runat="server" CssClass="text-Right-Blue" Text="客戶名稱" Width="120px" />
            <asp:TextBox ID="eCustName_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eCustName_Search_TextChanged" Width="400px" />
            <asp:GridView ID="gridCustData_Search" runat="server" CssClass="text-Left-Black" Width="600px" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="code" DataSourceID="dsCustData" ForeColor="#333333" GridLines="None" >
                <AlternatingRowStyle BackColor="White" />
                <Columns>
                    <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                    <asp:BoundField DataField="code" HeaderText="客戶編號" ReadOnly="True" SortExpression="code" />
                    <asp:BoundField DataField="name" HeaderText="客戶名稱" SortExpression="name" />
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
    <asp:SqlDataSource ID="dsCustData" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT code, name FROM CUSTOM WHERE (types = 'C') AND (name LIKE @Name)">
        <SelectParameters>
            <asp:Parameter Name="Name" />
        </SelectParameters>
    </asp:SqlDataSource>
</body>
</html>
