<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CarFixWorkHistoryP.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CarFixWorkHistoryP" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="CarFixWorkHistoryPForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">車輛資歷卡</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbCarID_Search" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eCarID_Search" runat="server" CssClass="text-Left-Black" OnTextChanged="eCarID_Search_TextChanged" AutoPostBack="true" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbBuDate_Search" runat="server" CssClass="text-Right-Blue" Text="開單日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eBudate_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_Budate_Search" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eBudate_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbEXCEL" runat="server" CssClass="button-Black" Text="匯出 EXCEL" OnClick="bbEXCEL_Click" Visible="false" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbPreview" runat="server" CssClass="button-Blue" Text="預覽報表" OnClick="bbPreview_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
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
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridCarFixWorkHistory_P" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1"
            DataSourceID="sdsCarFixWorkHistory_P" GridLines="None" PageSize="5" Width="100%">
            <Columns>
                <asp:BoundField DataField="CAR_ID" HeaderText="車號" SortExpression="CAR_ID" />
                <asp:BoundField DataField="BUDATE" HeaderText="開單日" SortExpression="BUDATE" DataFormatString="{0:}" />
                <asp:BoundField DataField="BUMAN" HeaderText="BUMAN" SortExpression="BUMAN" Visible="false" />
                <asp:BoundField DataField="BuMan_C" HeaderText="維修員 (代表)" ReadOnly="True" SortExpression="BuMan_C" />
                <asp:BoundField DataField="CompanyNo" HeaderText="CompanyNo" SortExpression="CompanyNo" ReadOnly="True" Visible="false" />
                <asp:BoundField DataField="CompanyNo_C" HeaderText="目前配置單位" ReadOnly="True" SortExpression="CompanyNo_C" />
                <asp:BoundField DataField="DEPNO" HeaderText="DEPNO" SortExpression="DEPNO" Visible="false" />
                <asp:BoundField DataField="DepNo_C" HeaderText="維修時配置單位" ReadOnly="True" SortExpression="DepNo_C" />
                <asp:BoundField DataField="service" HeaderText="service" SortExpression="service" Visible="false" />
                <asp:BoundField DataField="Service_C" HeaderText="維修代號" SortExpression="Service_C" ReadOnly="True" />
                <asp:BoundField DataField="DRIVER" HeaderText="DRIVER" SortExpression="DRIVER" Visible="false" />
                <asp:BoundField DataField="Driver_C" HeaderText="駕駛員" ReadOnly="True" SortExpression="Driver_C" />
                <asp:BoundField DataField="REMARK" HeaderText="維修項目 / 主要零件" SortExpression="REMARK" />
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
    <asp:SqlDataSource ID="sdsCarFixWorkHistory_P" runat="server"
        ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAR_ID, BUDATE, BUMAN, (SELECT NAME FROM Employee WHERE (EMPNO = a.BUMAN)) AS BuMan_C, (SELECT CompanyNo FROM Car_InfoA WHERE (Car_ID = a.CAR_ID)) AS CompanyNo, (SELECT NAME FROM DepartMent WHERE (DEPNO = (SELECT CompanyNo FROM Car_InfoA AS Car_InfoA_1 WHERE (Car_ID = a.CAR_ID)))) AS CompanyNo_C, DEPNO, (SELECT NAME FROM Department AS Department_1 WHERE (DEPNO = a.DEPNO)) AS DepNo_C, service, (SELECT DESCRIPTION FROM DBDICB WHERE (CLASSNO = a.service) AND (FKEY = '工作單A         FixworkA        SERVICE')) AS Service_C, DRIVER, (SELECT NAME FROM EMployee AS EMployee_1 WHERE (EMPNO = a.DRIVER)) AS Driver_C, REMARK FROM FixWorkA AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>

    <asp:Panel ID="plReport" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCloseReport" runat="server" CssClass="button-Red" Text="結束預覽" OnClick="bbCloseReport_Click" Width="120px" />
        <rsweb:ReportViewer ID="rvPrint" runat="server" BackColor="" ClientIDMode="AutoID" HighlightBackgroundColor=""
            InternalBorderColor="204, 204, 204" InternalBorderStyle="Solid" InternalBorderWidth="1px"
            LinkActiveColor="" LinkActiveHoverColor="" LinkDisabledColor=""
            PrimaryButtonBackgroundColor="" PrimaryButtonForegroundColor="" PrimaryButtonHoverBackgroundColor="" PrimaryButtonHoverForegroundColor=""
            SecondaryButtonBackgroundColor="" SecondaryButtonForegroundColor="" SecondaryButtonHoverBackgroundColor="" SecondaryButtonHoverForegroundColor=""
            SplitterBackColor=""
            ToolbarDividerColor="" ToolbarForegroundColor="" ToolbarForegroundDisabledColor=""
            ToolbarHoverBackgroundColor="" ToolbarHoverForegroundColor=""
            ToolBarItemBorderColor="" ToolBarItemBorderStyle="Solid" ToolBarItemBorderWidth="1px" ToolBarItemHoverBackColor=""
            ToolBarItemPressedBorderColor="51, 102, 153" ToolBarItemPressedBorderStyle="Solid" ToolBarItemPressedBorderWidth="1px"
            ToolBarItemPressedHoverBackColor="153, 187, 226" Height="600px" Width="100%" PageCountMode="Actual">
            <LocalReport ReportPath="Report\CarFixWorkHistory.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
