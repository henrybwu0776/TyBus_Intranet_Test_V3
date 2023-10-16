<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="InputDate.aspx.cs" Inherits="TyBus_Intranet_Test_V3.InputDate" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link rel="stylesheet" type="text/css" href="Style/MainStyle.css" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>點選日期</title>
</head>
<body>
    <form id="InputDateForm" runat="server">
        <div>
            <asp:Calendar ID="ChoiseDate" runat="server" BackColor="White" BorderColor="Black" BorderStyle="Solid" CellSpacing="1" Font-Names="Verdana" Font-Size="9pt" ForeColor="Black" Height="250px" NextPrevFormat="ShortMonth" Width="330px">
                <DayHeaderStyle Font-Bold="True" Font-Size="8pt" ForeColor="#333333" Height="8pt" />
                <DayStyle BackColor="#CCCCCC" />
                <NextPrevStyle Font-Bold="True" Font-Size="8pt" ForeColor="White" />
                <OtherMonthDayStyle ForeColor="#999999" />
                <SelectedDayStyle BackColor="#333399" ForeColor="White" />
                <TitleStyle BackColor="#333399" BorderStyle="Solid" Font-Bold="True" Font-Size="12pt" ForeColor="White" Height="12pt" />
                <TodayDayStyle BackColor="#999999" ForeColor="White" />
            </asp:Calendar>
            <asp:Button ID="bbOK" runat="server" Text="OK" CssClass="button-Blue" Width="100px" OnClick="bbOK_Click" />
            <asp:Button ID="bbCancel" runat="server" Text="Cancel" CssClass="button-Red" Width="100px" OnClick="bbCancel_Click" />
        </div>
    </form>
</body>
</html>
