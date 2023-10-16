<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="Over65DriverList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.Over65DriverList" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="Over65DriverListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">年滿65歲駕駛員名冊</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalBaseDate" runat="server" CssClass="text-Right-Blue" Text="基準日：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCalBaseDate" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Search_TextChanged" Width="30%" />
                    <asp:Label ID="eDepName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Black" Text="印表" OnClick="bbPrint_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="離開" OnClick="bbClose_Click" Width="90%" />
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
        <asp:GridView ID="gridShowData" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#999999" BorderStyle="None" BorderWidth="1px" CellPadding="3" DataKeyNames="EMPNO" DataSourceID="sdsShowData" GridLines="Vertical" Width="100%">
            <AlternatingRowStyle BackColor="#DCDCDC" />
            <Columns>
                <asp:BoundField DataField="DEPNO" HeaderText="站別代號" SortExpression="DEPNO" />
                <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="EMPNO" HeaderText="工號" ReadOnly="True" SortExpression="EMPNO" />
                <asp:BoundField DataField="NAME" HeaderText="姓名" SortExpression="NAME" />
                <asp:BoundField DataField="BIRTHDAY" DataFormatString="{0:D}" HeaderText="出生日期" SortExpression="BIRTHDAY" />
                <asp:BoundField DataField="IDCARDNO" HeaderText="身分證號" SortExpression="IDCARDNO" />
                <asp:BoundField DataField="LicenceCheck" DataFormatString="{0:D}" HeaderText="駕照審驗期限" SortExpression="LicenceCheck" />
                <asp:BoundField DataField="BBCALL" DataFormatString="{0:D}" HeaderText="定期訓練審驗期限" SortExpression="BBCALL" />
            </Columns>
            <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
            <HeaderStyle BackColor="#000084" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <RowStyle BackColor="#EEEEEE" ForeColor="Black" />
            <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F1F1F1" />
            <SortedAscendingHeaderStyle BackColor="#0000A9" />
            <SortedDescendingCellStyle BackColor="#CAC9C9" />
            <SortedDescendingHeaderStyle BackColor="#000065" />
        </asp:GridView>
        <asp:SqlDataSource ID="sdsShowData" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DEPNO)) AS DepName, EMPNO, NAME, BIRTHDAY, IDCARDNO, LicenceCheck, BBCALL FROM EMPLOYEE AS e WHERE (ISNULL(EMPNO, '') = '')"></asp:SqlDataSource>
    </asp:Panel>
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCloseReport" runat="server" CssClass="button-Red" Text="關閉報表" OnClick="bbCloseReport_Click" Width="100px" />
        <rsweb:ReportViewer ID="rvPrint" runat="server" BackColor="" ClientIDMode="AutoID" HighlightBackgroundColor=""
            InternalBorderColor="204, 204, 204" InternalBorderStyle="Solid" InternalBorderWidth="1px" LinkActiveColor=""
            LinkActiveHoverColor="" LinkDisabledColor=""
            PrimaryButtonBackgroundColor="" PrimaryButtonForegroundColor="" PrimaryButtonHoverBackgroundColor="" PrimaryButtonHoverForegroundColor=""
            SecondaryButtonBackgroundColor="" SecondaryButtonForegroundColor="" SecondaryButtonHoverBackgroundColor="" SecondaryButtonHoverForegroundColor=""
            SplitterBackColor="" ToolbarDividerColor="" ToolbarForegroundColor="" ToolbarForegroundDisabledColor="" ToolbarHoverBackgroundColor=""
            ToolbarHoverForegroundColor="" ToolBarItemBorderColor="" ToolBarItemBorderStyle="Solid" ToolBarItemBorderWidth="1px" ToolBarItemHoverBackColor=""
            ToolBarItemPressedBorderColor="51, 102, 153" ToolBarItemPressedBorderStyle="Solid" ToolBarItemPressedBorderWidth="1px" ToolBarItemPressedHoverBackColor="153, 187, 226"
            Height="600px" Width="100%">
            <LocalReport ReportPath="Report\Over65DriverListP.rdlc">
                <DataSources>
                </DataSources>
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
