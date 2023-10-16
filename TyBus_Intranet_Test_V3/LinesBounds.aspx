<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="LinesBounds.aspx.cs" Inherits="TyBus_Intranet_Test_V3.LinesBounds" %>

<asp:Content ID="LinesBoundsForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">路線電子票證營收及優惠補助資料</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbImportData" runat="server" CssClass="text-Right-Blue" Text="匯入資料" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlSourceTyoe" runat="server" CssClass="text-Left-Blue" Width="100%">
                        <asp:ListItem Value="00" Text="" Selected="True" />
                        <asp:ListItem Value="01" Text="各路線現金收入" Enabled="false" />
                        <asp:ListItem Value="02" Text="各路線刷卡收入" Enabled="false" />
                        <asp:ListItem Value="03" Text="表一資料" />
                        <asp:ListItem Value="04" Text="表三資料" />
                        <asp:ListItem Value="05" Text="學生卡" />
                        <asp:ListItem Value="06" Text="敬老及愛心卡" />
                        <asp:ListItem Value="07" Text="國道及台鐵轉乘_公路" />
                        <asp:ListItem Value="08" Text="國道及台鐵轉乘_市區" />
                        <asp:ListItem Value="09" Text="國小升國中半價補貼" />
                        <asp:ListItem Value="10" Text="新北公車間轉乘補貼" />
                        <asp:ListItem Value="11" Text="新北公車票差補貼" />
                        <asp:ListItem Value="12" Text="國道準運價" />
                        <asp:ListItem Value="13" Text="北捷轉乘補貼" />
                        <asp:ListItem Value="14" Text="1280 暫撥款" />
                        <asp:ListItem Value="15" Text="北捷定期票預撥款" />
                        <asp:ListItem Value="16" Text="新北老人補貼" />
                        <asp:ListItem Value="17" Text="台灣好行" />
                    </asp:DropDownList>
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:Label ID="lbYearTitle" runat="server" CssClass="text-Left-Blue" Text="中華民國 " />
                    <asp:TextBox ID="eImportYear" runat="server" CssClass="text-Left-Black" Width="20%" />
                    <asp:Label ID="lbYear" runat="server" CssClass="text-Left-Blue" Text=" 年 " />
                    <asp:DropDownList ID="ddlImportMonth" runat="server" CssClass="text-Left-Black" Width="35%">
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
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:Label ID="lbTicketPrice" runat="server" CssClass="text-Left-Blue" Text="全票票價_市區：" />
                    <asp:TextBox ID="eTicketPrice_City" runat="server" CssClass="text-Left-Black" Width="50px" />
                    <asp:Label ID="lbTicketPrice_2" runat="server" CssClass="text-Left-Blue" Text="  國道：" />
                    <asp:TextBox ID="eTicketPrice_Highway" runat="server" CssClass="text-Left-Black" Width="50px" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="5">
                    <asp:FileUpload ID="fuExcel" runat="server" CssClass="text-Left-Black" Width="97%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbImportData" runat="server" CssClass="button-Blue" Text="匯入" OnClick="bbImportData_Click" Width="100%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearchData" runat="server" CssClass="button-Red" Text="查詢" OnClick="bbSearchData_Click" Width="100%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExportExcel" runat="server" CssClass="button-Black" Text="匯出成本試算" OnClick="bbExportExcel_Click" Width="100%" />
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
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridShowData" runat="server" AllowPaging="True" AutoGenerateColumns="False" DataKeyNames="IndexNo" DataSourceID="dsImportData" Width="100%" BackColor="White" BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" CellPadding="3" GridLines="Horizontal">
            <AlternatingRowStyle BackColor="#F7F7F7" />
            <Columns>
                <asp:BoundField DataField="IndexNo" HeaderText="IndexNo" ReadOnly="True" SortExpression="IndexNo" Visible="False" />
                <asp:BoundField DataField="CalYM" HeaderText="資料年月" SortExpression="CalYM" />
                <asp:BoundField DataField="TicketType" HeaderText="TicketType" SortExpression="TicketType" Visible="False" />
                <asp:BoundField DataField="TicketLineNo" HeaderText="路線IC代碼" SortExpression="TicketLineNo" />
                <asp:BoundField DataField="LinesGovNo" HeaderText="路線編號" SortExpression="LinesGovNo" />
                <asp:BoundField DataField="TicketType_C" HeaderText="票種名稱" ReadOnly="True" SortExpression="TicketType_C" />
                <asp:BoundField DataField="TicketPCount" HeaderText="刷卡量(次)" SortExpression="TicketPCount" />
                <asp:BoundField DataField="TotalAmount" HeaderText="應收總金額" SortExpression="TotalAmount" />
                <asp:BoundField DataField="RealIncome" HeaderText="實際營收" SortExpression="RealIncome" />
                <asp:BoundField DataField="Subsidy_8KM" HeaderText="基本里程補貼" SortExpression="Subsidy_8KM" />
                <asp:BoundField DataField="Subsidy_Limited" HeaderText="票價上限補貼" SortExpression="Subsidy_Limited" />
                <asp:BoundField DataField="Subsidy_Student" HeaderText="學生卡補貼" SortExpression="Subsidy_Student" />
                <asp:BoundField DataField="Subsidy_Social" HeaderText="社福補助" SortExpression="Subsidy_Social" />
                <asp:BoundField DataField="Subsidy_MRT" HeaderText="捷運轉乘補貼" SortExpression="Subsidy_MRT" />
                <asp:BoundField DataField="Subsidy_Area" HeaderText="A21-A23區間補貼" SortExpression="Subsidy_Area" />
                <asp:BoundField DataField="Subsidy_Highway" HeaderText="國道台鐵轉乘補貼" SortExpression="Subsidy_Highway" />
                <asp:BoundField DataField="Subsidy_Others" HeaderText="其他補貼" SortExpression="Subsidy_Others" />
                <asp:BoundField DataField="Subsidy_Dift" HeaderText="法定差額補貼" SortExpression="Subsidy_Dift" />
                <asp:BoundField DataField="BuMan" HeaderText="BuMan" SortExpression="BuMan" Visible="False" />
                <asp:BoundField DataField="BuDate" HeaderText="BuDate" SortExpression="BuDate" Visible="False" />
                <asp:BoundField DataField="ModifyMan" HeaderText="ModifyMan" SortExpression="ModifyMan" Visible="False" />
                <asp:BoundField DataField="ModifyDate" HeaderText="ModifyDate" SortExpression="ModifyDate" Visible="False" />
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
    <asp:SqlDataSource ID="dsImportData" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT IndexNo, CalYM, TicketType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '路線優惠補助    fmLinesBounds   TicketType') AND (CLASSNO = a.TicketType)) AS TicketType_C, TicketLineNo, LinesGovNo, TicketPCount, TotalAmount, RealIncome, Subsidy_8KM, Subsidy_Limited, Subsidy_Student, Subsidy_Social, Subsidy_MRT, Subsidy_Area, Subsidy_Highway, Subsidy_Others, Subsidy_Dift, BuMan, BuDate, ModifyMan, ModifyDate FROM LinesBounds AS a WHERE (ISNULL(IndexNo, '') = '')"></asp:SqlDataSource>
</asp:Content>
