<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ImportTLD_Data.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ImportTLD_Data" %>
<asp:Content ID="ImportTLD_DataForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">匯入交易明細檔資料</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="3">
                    <asp:RadioButtonList ID="eImportSourceType" runat="server" CssClass="text-Left-Black" Width="95%" RepeatColumns="3">
                        <asp:ListItem Value="TLD_Data" Text="悠遊卡 TLD 交易明細" />
                        <asp:ListItem Value="TLD_Data2" Text="悠遊卡 TLD 交易明細 (有錯誤代碼)" />
                        <asp:ListItem Value="ACER_Data" Text="宏碁清分記錄" />
                    </asp:RadioButtonList>
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbFileName" runat="server" CssClass="text-Right-Blue" Text="交易明細檔" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:FileUpload ID="fuFileName" runat="server" CssClass="text-Left-Black" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Label ID="lbClearDate" runat="server" CssClass="text-Right-Blue" Text="清分日期" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eClearDate" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Blue" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Button ID="bbExport" runat="server" CssClass="button-Black" Text="匯出EXCEL" OnClick="bbExport_Click" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Button ID="bbImport" runat="server" CssClass="button-Blue" Text="資料轉入" OnClick="bbImport_Click" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" OnClick="bbExit_Click" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gvTLD_Data" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataSourceID="dsTLD_Data" ForeColor="#333333" GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:BoundField DataField="TitleMark" HeaderText="檔頭標籤" SortExpression="TitleMark" Visible="False" />
                <asp:BoundField DataField="DataType" HeaderText="資料屬性" SortExpression="DataType" Visible="False" />
                <asp:BoundField DataField="TradeMode" HeaderText="交易模式" SortExpression="TradeMode" />
                <asp:BoundField DataField="ClearDate" HeaderText="清分日期" SortExpression="ClearDate" />
                <asp:BoundField DataField="ServiceDate" HeaderText="營運日期" SortExpression="ServiceDate" />
                <asp:BoundField DataField="StationCode" HeaderText="站別代碼" SortExpression="StationCode" />
                <asp:BoundField DataField="MachineCode" HeaderText="設備代碼" SortExpression="MachineCode" />
                <asp:BoundField DataField="Car_ID" HeaderText="車號" SortExpression="Car_ID" />
                <asp:BoundField DataField="LinesNo" HeaderText="路線代碼" SortExpression="LinesNo" />
                <asp:BoundField DataField="DriverNo" HeaderText="司機代碼" SortExpression="DriverNo" />
                <asp:BoundField DataField="StartTime" HeaderText="開班時間" SortExpression="StartTime" />
                <asp:BoundField DataField="IDType" HeaderText="身份別" SortExpression="IDType" />
                <asp:BoundField DataField="ChipCardNo" HeaderText="晶片卡號" SortExpression="ChipCardNo" />
                <asp:BoundField DataField="SP_Card" HeaderText="特種票種" SortExpression="SP_Card" />
                <asp:BoundField DataField="TradeDate" HeaderText="交易日期" SortExpression="TradeDate" />
                <asp:BoundField DataField="TradeCode" HeaderText="交易序號" SortExpression="TradeCode" />
                <asp:BoundField DataField="Amount" HeaderText="交易金額" SortExpression="Amount" />
                <asp:BoundField DataField="TransDiscount" HeaderText="轉乘優惠" SortExpression="TransDiscount" />
                <asp:BoundField DataField="PersonalDiscount" HeaderText="個人優惠" SortExpression="PersonalDiscount" />
                <asp:BoundField DataField="LeaveStation" HeaderText="出站" SortExpression="LeaveStation" />
                <asp:BoundField DataField="In_OutFlag" HeaderText="上下車" SortExpression="In_OutFlag" />
                <asp:BoundField DataField="AreaCode" HeaderText="區碼" SortExpression="AreaCode" />
                <asp:BoundField DataField="InStation" HeaderText="進站" SortExpression="InStation" />
                <asp:BoundField DataField="ErrorCode" HeaderText="錯誤代碼" SortExpression="ErrorCode" />
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
        <asp:GridView ID="gvACER_Data" runat="server" CellPadding="4" ForeColor="#333333" GridLines="None" AllowPaging="True" AutoGenerateColumns="False" DataKeyNames="IndexNo" DataSourceID="dsACER_Data">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <Columns>
                <asp:BoundField DataField="IndexNo" HeaderText="序號" InsertVisible="False" ReadOnly="True" SortExpression="IndexNo" Visible="False" />
                <asp:BoundField DataField="TicketType" HeaderText="票證公司" SortExpression="TicketType" />
                <asp:BoundField DataField="ClearDate" HeaderText="清分日期" SortExpression="ClearDate" />
                <asp:BoundField DataField="ServiceDate" HeaderText="營運日期" SortExpression="ServiceDate" />
                <asp:BoundField DataField="StationCode" HeaderText="站別代碼" SortExpression="StationCode" />
                <asp:BoundField DataField="Car_ID" HeaderText="車號" SortExpression="Car_ID" />
                <asp:BoundField DataField="DriverNo" HeaderText="司機代碼" SortExpression="DriverNo" />
                <asp:BoundField DataField="LinesNo" HeaderText="路線代碼" SortExpression="LinesNo" />
                <asp:BoundField DataField="TSAmount" HeaderText="總營運金額" SortExpression="TSAmount" />
                <asp:BoundField DataField="StudentAmount" HeaderText="學生卡金額" SortExpression="StudentAmount" />
                <asp:BoundField DataField="OldAmount_N" HeaderText="敬老愛心卡金額" SortExpression="OldAmount_N" />
                <asp:BoundField DataField="NormalAmount_8KM" HeaderText="普卡金額 (8KM)" SortExpression="NormalAmount_8KM" />
                <asp:BoundField DataField="TotalAmount" HeaderText="總合計" SortExpression="TotalAmount" />
                <asp:BoundField DataField="TotalPassCount" HeaderText="總載客人數" SortExpression="TotalPassCount" />
                <asp:BoundField DataField="PassCount_N" HeaderText="普卡載客人數" SortExpression="PassCount_N" />
                <asp:BoundField DataField="PassCount_U" HeaderText="聯營卡載客人數" SortExpression="PassCount_U" />
                <asp:BoundField DataField="PassCount_S" HeaderText="學生卡載客人數" SortExpression="PassCount_S" />
                <asp:BoundField DataField="PassCount_O" HeaderText="敬老愛心卡載客人數" SortExpression="PassCount_O" />
                <asp:BoundField DataField="TPassCount" HeaderText="基北北桃月票人數" SortExpression="TPassCount" />
                <asp:BoundField DataField="TPassAmount" HeaderText="基北北桃月票金額" SortExpression="TPassAmount" />
            </Columns>
            <EditRowStyle BackColor="#999999" />
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#E9E7E2" />
            <SortedAscendingHeaderStyle BackColor="#506C8C" />
            <SortedDescendingCellStyle BackColor="#FFFDF8" />
            <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
        </asp:GridView>
    </asp:Panel>
    <asp:SqlDataSource ID="dsTLD_Data" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT [TitleMark], [DataType], [TradeMode], [ClearDate], [ServiceDate], [StationCode], [MachineCode], [Car_ID], [LinesNo], [DriverNo], [StartTime], [IDType], [ChipCardNo], [SP_Card], [TradeDate], [TradeCode], [Amount], [TransDiscount], [PersonalDiscount], [LeaveStation], [In_OutFlag], [AreaCode], [InStation], [ErrorCode] FROM [TLD_Data]"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsACER_Data" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT IndexNo, TicketType, ClearDate, ServiceDate, StationCode, Car_ID, DriverNo, LinesNo, TSAmount, StudentAmount, OldAmount_N, NormalAmount_8KM, TotalAmount, TotalPassCount, PassCount_N, PassCount_U, PassCount_S, PassCount_O, TPassCount, TPassAmount FROM ACER_Data"></asp:SqlDataSource>
</asp:Content>
