<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="SearchSup.aspx.cs" Inherits="TyBus_Intranet_Test_V3.SearchSup" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link rel="stylesheet" type="text/css" href="Style/MainStyle.css" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>查詢廠商名稱</title>
</head>
<body>
    <form id="SearchSupForm" runat="server">
        <div>
            <asp:Label ID="lbSupName_Search" runat="server" CssClass="text-Right-Blue" Text="廠商名稱" Width="120px" />
            <asp:TextBox ID="eSupName_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eSupName_Search_TextChanged" Width="400px" />
            <asp:GridView ID="gridSupData_Search" runat="server" CssClass="text-Left-Black" Width="600px" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="code" ForeColor="#333333" GridLines="None">
                <AlternatingRowStyle BackColor="White" />
                <Columns>
                    <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                    <asp:BoundField DataField="code" HeaderText="廠商編號" ReadOnly="True" SortExpression="code" />
                    <asp:BoundField DataField="name" HeaderText="廠商名稱" SortExpression="name" />
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
</body>
</html>
