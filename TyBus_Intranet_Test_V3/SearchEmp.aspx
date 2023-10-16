<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="SearchEmp.aspx.cs" Inherits="TyBus_Intranet_Test_V3.SearchEmp" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link rel="stylesheet" type="text/css" href="Style/MainStyle.css" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>查詢員工工號</title>
</head>
<body>
    <form id="SearchCustForm" runat="server">
        <div>
            <asp:Label ID="lbEmpName_Search" runat="server" CssClass="text-Right-Blue" Text="員工姓名：" Width="120px" />
            <asp:TextBox ID="eEmpName_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="True" OnTextChanged="eEmpName_Search_TextChanged" Width="400px" />
            <asp:GridView ID="gridEmpData_Search" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="EMPNO" ForeColor="#333333" GridLines="None" Width="600px">
                <AlternatingRowStyle BackColor="White" />
                <Columns>
                    <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                    <asp:BoundField DataField="EMPNO" HeaderText="員工工號" ReadOnly="True" SortExpression="EMPNO" />
                    <asp:BoundField DataField="NAME" HeaderText="姓名" SortExpression="NAME" />
                    <asp:BoundField DataField="worktype" HeaderText="在職狀況" SortExpression="worktype" />
                    <asp:BoundField DataField="Leaveday" HeaderText="離職日期" ReadOnly="True" SortExpression="Leaveday" />
                </Columns>
                <EditRowStyle BackColor="#2461BF" />
                <EmptyDataTemplate>
                    <asp:Label ID="lbEmpty" runat="server" CssClass="text-Left-Red" Text="查無資料"></asp:Label>
                </EmptyDataTemplate>
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
