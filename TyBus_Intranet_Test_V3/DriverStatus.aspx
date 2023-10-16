<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DriverStatus.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DriverStatus" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
<asp:Content ID="DriverStatusForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">駕駛員月出勤及請休假狀況表</a>
    </div>
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbCalYM_Search" runat="server" CssClass="text-Right-Blue" Text="計算年月" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eCalYear_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbCalYear_Search" runat="server" CssClass="text-Left-Blue" Text=" 年 " />
                    <asp:DropDownList ID="ddlCalMonth_Search" runat="server" CssClass="text-Left-Black" Width="30%">
                        <asp:ListItem Value="01" Text="一月" />
                        <asp:ListItem Value="02" Text="二月" />
                        <asp:ListItem Value="03" Text="三月" />
                        <asp:ListItem Value="04" Text="四月" />
                        <asp:ListItem Value="05" Text="五月" />
                        <asp:ListItem Value="06" Text="六月" />
                        <asp:ListItem Value="07" Text="七月" />
                        <asp:ListItem Value="08" Text="八月" />
                        <asp:ListItem Value="09" Text="九月" />
                        <asp:ListItem Value="10" Text="十月" />
                        <asp:ListItem Value="11" Text="十一月" />
                        <asp:ListItem Value="12" Text="十二月" />
                    </asp:DropDownList>
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDriver_Search" runat="server" CssClass="text-Right-Blue" Text="員工工號" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eDriver_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" Width="30%" />
                    <asp:Label ID="eDriverName_Search" runat="server" CssClass="text-Left-Blue" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" Width="95%" OnClick="bbSearch_Click" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出EXCEL" Width="95%" OnClick="bbExcel_Click" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="離開" Width="95%" OnClick="bbClose_Click" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridDriverStatus" runat="server" BackColor="White" BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" CellPadding="3" GridLines="Horizontal" Width="100%" AutoGenerateColumns="False" DataSourceID="dsDriverStatus">
            <AlternatingRowStyle BackColor="#F7F7F7" />
            <Columns>
                <asp:BoundField DataField="DT" DataFormatString="{0:D}" ReadOnly="True" SortExpression="DT" />
                <asp:BoundField DataField="Name" HeaderText="姓名" ReadOnly="True" SortExpression="Name" />
                <asp:BoundField DataField="Birthday" DataFormatString="{0:D}" HeaderText="出生日期" ReadOnly="True" SortExpression="Birthday" />
                <asp:BoundField DataField="IDCardNo" HeaderText="身分證字號" ReadOnly="True" SortExpression="IDCardNo" />
                <asp:BoundField DataField="WorkType" HeaderText="在職情形" ReadOnly="True" SortExpression="WorkType" />
                <asp:BoundField DataField="YearCount" HeaderText="年資" ReadOnly="True" SortExpression="YearCount" />
                <asp:BoundField DataField="Assumeday" DataFormatString="{0:D}" HeaderText="受雇日期" ReadOnly="True" SortExpression="Assumeday" />
                <asp:BoundField DataField="DepNo" HeaderText="部門" ReadOnly="True" SortExpression="DepNo" />
                <asp:BoundField DataField="EmpNo" HeaderText="工號" ReadOnly="True" SortExpression="EmpNo" />
                <asp:BoundField DataField="TelNo1" HeaderText="電話" ReadOnly="True" SortExpression="TelNo1" />
                <asp:BoundField DataField="CellPhone" HeaderText="行動電話" ReadOnly="True" SortExpression="CellPhone" />
                <asp:BoundField DataField="Addr1" HeaderText="地址" ReadOnly="True" SortExpression="Addr1" />
                <asp:BoundField DataField="Title" HeaderText="職稱" ReadOnly="True" SortExpression="Title" />
                <asp:BoundField DataField="WorkHR" HeaderText="每日工時" ReadOnly="True" SortExpression="WorkHR" />
                <asp:BoundField DataField="WorkState" HeaderText="每日排班" ReadOnly="True" SortExpression="WorkState" />
                <asp:BoundField DataField="ESCType_04" HeaderText="事假紀錄" ReadOnly="True" SortExpression="ESCType_04" />
                <asp:BoundField DataField="ESCType_05" HeaderText="病假紀錄" ReadOnly="True" SortExpression="ESCType_05" />
                <asp:BoundField DataField="MonthPay" HeaderText="當月薪資" ReadOnly="True" SortExpression="MonthPay" />
            </Columns>
            <FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
            <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#F7F7F7" />
            <PagerStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" HorizontalAlign="Right" />
            <RowStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="#F7F7F7" />
            <SortedAscendingCellStyle BackColor="#F4F4FD" />
            <SortedAscendingHeaderStyle BackColor="#5A4C9D" />
            <SortedDescendingCellStyle BackColor="#D8D8F0" />
            <SortedDescendingHeaderStyle BackColor="#3E3277" />
        </asp:GridView>
    </asp:Panel>
    <asp:SqlDataSource ID="dsDriverStatus" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select cast(null as datetime) as DT, cast(null as varchar) as [Name], cast(null as datetime) as Birthday, cast(null as varchar) as IDCardNo, 
       cast(null as varchar) as WorkType, cast(null as float) as YearCount, cast(null as datetime) as Assumeday, cast(null as varchar) as DepNo, 
	   cast(null as varchar) as EmpNo, cast(null as varchar) as TelNo1, cast(null as varchar) as Addr1, 
	   cast(null as varchar) as Title, cast(null as int) as WorkHR, cast(null as varchar) as WorkState, cast(null as varchar) as ESCType_04, 
	   cast(null as varchar) as ESCType_05, cast(null as varchar) as MonthPay, cast(null as varchar) as CellPhone "></asp:SqlDataSource>
</asp:Content>
