<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmpOverList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmpOverList" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="EmpOverListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">加班資料明細表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="單位" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_S" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_E" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbEmpNo" runat="server" CssClass="text-Right-Blue" Text="員工工號" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eEmpNo" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbRealDay" runat="server" CssClass="text-Right-Blue" Text="加班日期" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eRealDay_S" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eRealDay_E" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPreview" runat="server" CssClass="button-Black" Text="預覽" OnClick="bbPreview_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
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
        <asp:GridView ID="gridEmpOverList" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataSourceID="sdsEmpOverList" GridLines="None" PageSize="20" Width="100%">
            <Columns>
                <asp:BoundField DataField="applyman" HeaderText="員工編號" SortExpression="applyman" />
                <asp:BoundField DataField="ApplyName" HeaderText="員工姓名" ReadOnly="True" SortExpression="ApplyName" />
                <asp:BoundField DataField="depno" HeaderText="depno" SortExpression="depno" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="部門" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="realday" DataFormatString="{0:d}" HeaderText="加班日期" SortExpression="realday" />
                <asp:BoundField DataField="applytype" HeaderText="applytype" SortExpression="applytype" Visible="False" />
                <asp:BoundField DataField="ApplyType_C" HeaderText="加班類別" ReadOnly="True" SortExpression="ApplyType_C" />
                <asp:BoundField DataField="hours" HeaderText="加班時數" SortExpression="hours" />
                <asp:BoundField DataField="Over100" HeaderText="1倍時數" ReadOnly="True" SortExpression="Over100" />
                <asp:BoundField DataField="Over133" HeaderText="1.33時數" ReadOnly="True" SortExpression="Over133" />
                <asp:BoundField DataField="Over166" HeaderText="1.66時數" ReadOnly="True" SortExpression="Over166" />
                <asp:BoundField DataField="Over266" HeaderText="2.66時數" ReadOnly="True" SortExpression="Over266" />
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
    <asp:Panel ID="plReport" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCloseReport" runat="server" CssClass="button-Red" OnClick="bbCloseReport_Click" Text="結束預覽" Width="120px" />
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
            <LocalReport ReportPath="Report\EmpOverListP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsEmpOverList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT applyman, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.applyman)) AS ApplyName, depno, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.depno)) AS DepName, realday, applytype, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.applytype) AND (FKEY = '加班資料檔      OVERDUTY        APPLYTYPE')) AS ApplyType_C, hours, CASE WHEN ApplyType IN ('02' , '04') THEN [Hours] - FeedNum - BackNum ELSE 0 END AS Over100, ISNULL(feednum, 0) + ISNULL(FRONTOVER2, 0) AS Over133, ISNULL(backnum, 0) + ISNULL(POSTOVER2, 0) AS Over166, ISNULL(POSTOVER22, 0) AS Over266 FROM OVERDUTY AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
