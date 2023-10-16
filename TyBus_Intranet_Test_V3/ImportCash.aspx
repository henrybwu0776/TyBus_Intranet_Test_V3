<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ImportCash.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ImportCash" %>

<asp:Content ID="ImportCashForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">投現收入轉檔</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Label ID="lbBuDate" runat="server" CssClass="text-Right-Blue" Text="繳銷日期：" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:TextBox ID="eBuDate" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Label ID="lbFilePath" runat="server" CssClass="text-Right-Blue" Text="資料檔路徑：" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col" colspan="2">
                    <asp:FileUpload ID="eFilePath" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPreview" runat="server" CssClass="button-Black" Text="預覽資料" OnClick="bbPreview_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
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
        <table class="TableSetting">
            <tr>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbImport" runat="server" CssClass="button-Black" Text="匯入ERP" OnClick="bbImport_Click" Width="90%" />
                </td>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbCloseData" runat="server" CssClass="button-Red" Text="結束預覽" OnClick="bbCloseData_Click" Width="90%" />
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
        <asp:GridView ID="gridShowData" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1"
            DataSourceID="sdsShowData" GridLines="None" Width="90%">
            <Columns>
                <asp:BoundField DataField="BuDate" DataFormatString="{0:d}" HeaderText="繳銷日" ReadOnly="True" SortExpression="BuDate" />
                <asp:BoundField DataField="DepNo" HeaderText="站別編號" ReadOnly="True" SortExpression="DepNo" />
                <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="Driver" HeaderText="駕駛員工號" ReadOnly="True" SortExpression="Driver" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員" ReadOnly="True" SortExpression="DriverName" />
                <asp:BoundField DataField="Car_No" HeaderText="車輛代碼" ReadOnly="True" SortExpression="Car_No" />
                <asp:BoundField DataField="Car_ID" HeaderText="牌照號碼" ReadOnly="True" SortExpression="Car_ID" />
                <asp:BoundField DataField="LinesNo" HeaderText="路線別" ReadOnly="True" SortExpression="LinesNo" />
                <asp:BoundField DataField="CashType" HeaderText="類別代碼" ReadOnly="True" SortExpression="CashType" />
                <asp:BoundField DataField="CashType_C" HeaderText="繳銷類別" ReadOnly="True" SortExpression="CashType_C" />
                <asp:BoundField DataField="Cash_1000" HeaderText="仟元鈔" ReadOnly="True" SortExpression="Cash_1000" />
                <asp:BoundField DataField="Cash_500" HeaderText="伍佰元鈔" ReadOnly="True" SortExpression="Cash_500" />
                <asp:BoundField DataField="Cash_100" HeaderText="百元鈔" ReadOnly="True" SortExpression="Cash_100" />
                <asp:BoundField DataField="CashAMT" HeaderText="納金額" ReadOnly="True" SortExpression="CashAMT" />
                <asp:BoundField DataField="BuMan" HeaderText="開單人工號" ReadOnly="True" SortExpression="BuMan" />
                <asp:BoundField DataField="BuMan_C" HeaderText="開單人" ReadOnly="True" SortExpression="BuMan_C" />
                <asp:BoundField DataField="Remark" HeaderText="備註" ReadOnly="True" SortExpression="Remark" />
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
    <asp:SqlDataSource ID="sdsShowData" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS DateTime) AS BuDate, CAST('' AS varchar) AS DepNo, CAST('' AS varchar) AS DepName, CAST('' AS varchar) AS Driver, CAST('' AS varchar) AS DriverName, CAST('' AS varchar) AS Car_No, CAST('' AS varchar) AS Car_ID, CAST('' AS varchar) AS LinesNo, CAST('' AS varchar) AS CashType, CAST('' AS varchar) AS CashType_C, CAST(0 AS int) AS Cash_1000, CAST(0 AS int) AS Cash_500, CAST(0 AS int) AS Cash_100, CAST(0 AS int) AS CashAMT, CAST('' AS varchar) AS BuMan, CAST('' AS varchar) AS BuMan_C, CAST('' AS varchar) AS Remark FROM CashBill WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
