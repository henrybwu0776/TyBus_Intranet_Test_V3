<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ViolationCaseP.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ViolationCaseP" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="ViolationCasePForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">違規案件統計表</a>
    </div>
    <br />
    <asp:Panel ID="plSearech" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="7">
                    <asp:RadioButtonList ID="rbPrintMode" runat="server" CssClass="text-Left-Black" RepeatColumns="5" Width="95%"
                        OnSelectedIndexChanged="rbPrintMode_SelectedIndexChanged">
                        <asp:ListItem Value="11" Text="交通違規報表(依站別)" />
                        <asp:ListItem Value="12" Text="交通違規報表(依駕駛)" />
                        <asp:ListItem Value="13" Text="行政裁罰報表(依站別)" />
                        <asp:ListItem Value="21" Text="交通違規扣點統計表" />
                        <asp:ListItem Value="99" Text="統計報表" Selected="True" />
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbBuildDate" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eBuildDate_Start" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eBuildDate_End" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbViolationDate" runat="server" CssClass="text-Right-Blue" Text="違規日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eViolationDate_Start" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit2" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eViolationDate_End" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColWidth-7Col"></td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="站別編號：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Start" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit3" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_End" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbDriver" runat="server" CssClass="text-Right-Blue" Text="違規駕駛編號：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eDriver" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbCar_ID" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eCar_ID" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbViolationRule_Search" runat="server" CssClass="text-Right-Blue" Text="違規法條：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="6">
                    <asp:TextBox ID="eViolationRule_Search" runat="server" CssClass="text-Left-Black" Width="97%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-7Col">
                    <asp:Button ID="bbPreviewData" runat="server" CssClass="button-Black" Text="預覽資料" OnClick="bbPreviewData_Click" Width="95%" />
                </td>
                <td class="ColWidth-7Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Red" Text="列印報表" OnClick="bbPrint_Click" Width="95%" Visible="false" />
                </td>
                <td class="ColWidth-7Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Red" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="95%" Visible="false" />
                </td>
                <td class="ColWidth-7Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
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
            <LocalReport ReportPath="Report\ViolationCaseP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:Panel ID="plShowData" runat="server" GroupingText="案件清單" Width="95%">
        <asp:GridView ID="gvShowData_DepNo" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White"
            BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="CaseNo"
            DataSourceID="sdsViolationCaseP_DepNo" GridLines="None" PageSize="5" Width="100%">
            <Columns>
                <asp:BoundField DataField="CaseNo" HeaderText="CaseNo" ReadOnly="True" SortExpression="CaseNo" Visible="False" />
                <asp:BoundField DataField="CaseType" HeaderText="CaseType" SortExpression="CaseType" Visible="False" />
                <asp:BoundField DataField="CaseType_C" HeaderText="案件類別" ReadOnly="True" SortExpression="CaseType_C" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="站別" SortExpression="DepName" />
                <asp:BoundField DataField="ViolationDate" DataFormatString="{0:d}" HeaderText="日期" SortExpression="ViolationDate" />
                <asp:BoundField DataField="Driver" HeaderText="Driver" SortExpression="Driver" Visible="False" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員姓名" SortExpression="DriverName" />
                <asp:BoundField DataField="ViolationRule" HeaderText="法條" SortExpression="ViolationRule" />
                <asp:BoundField DataField="ViolationPoint" HeaderText="記 (扣) 點" SortExpression="ViolationPoint" />
                <asp:BoundField DataField="ViolationNote" HeaderText="違規事項" SortExpression="ViolationNote" />
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
        <asp:GridView ID="gvShowData_Driver" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White"
            BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="CaseNo"
            DataSourceID="sdsViolationCaseP_Driver" GridLines="None" PageSize="5" Width="100%">
            <Columns>
                <asp:BoundField DataField="CaseNo" HeaderText="CaseNo" ReadOnly="True" SortExpression="CaseNo" Visible="False" />
                <asp:BoundField DataField="CaseType" HeaderText="CaseType" SortExpression="CaseType" Visible="False" />
                <asp:BoundField DataField="CaseType_C" HeaderText="案件類別" ReadOnly="True" SortExpression="CaseType_C" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" ReadOnly="true" SortExpression="DepNo" Visible="false" />
                <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="true" SortExpression="DepName" />
                <asp:BoundField DataField="Driver" HeaderText="Driver" SortExpression="Driver" Visible="False" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員姓名" SortExpression="DriverName" />
                <asp:BoundField DataField="Car_ID" HeaderText="車號" SortExpression="Car_ID" />
                <asp:BoundField DataField="ViolationDate" DataFormatString="{0:d}" HeaderText="日期" SortExpression="ViolationDate" />
                <asp:BoundField DataField="ViolationRule" HeaderText="法條" SortExpression="ViolationRule" />
                <asp:BoundField DataField="ViolationPoint" HeaderText="記 (扣) 點" SortExpression="ViolationPoint" />
                <asp:BoundField DataField="ViolationNote" HeaderText="違規事項" SortExpression="ViolationNote" Visible="false" />
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
        <asp:GridView ID="gvShowData_Point" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="CaseNo" DataSourceID="sdsViolationCaseP_Point" GridLines="None" Width="100%">
            <Columns>
                <asp:BoundField DataField="CaseNo" HeaderText="CaseNo" ReadOnly="True" SortExpression="CaseNo" Visible="False" />
                <asp:BoundField DataField="CaseType" HeaderText="CaseType" SortExpression="CaseType" Visible="False" />
                <asp:BoundField DataField="CaseType_C" HeaderText="類別" ReadOnly="True" SortExpression="CaseType_C" />
                <asp:BoundField DataField="DepNo" HeaderText="站別編號" SortExpression="DepNo" />
                <asp:BoundField DataField="DepName" HeaderText="站別" SortExpression="DepName" />
                <asp:BoundField DataField="ViolationDate" DataFormatString="{0:D}" HeaderText="違規日期" SortExpression="ViolationDate" />
                <asp:BoundField DataField="DriverName" HeaderText="違規駕駛" SortExpression="DriverName" />
                <asp:BoundField DataField="ViolationRule" HeaderText="違規法條" SortExpression="ViolationRule" />
                <asp:BoundField DataField="ViolationPoint" HeaderText="記(扣)點" SortExpression="ViolationPoint" />
                <asp:BoundField DataField="ViolationNote" HeaderText="違規事項說明" SortExpression="ViolationNote" />
                <asp:BoundField DataField="IsAffectivePoint" HeaderText="有效扣點" ReadOnly="True" SortExpression="IsAffectivePoint" />
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
        <asp:GridView ID="gvShowData_99" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="White"
            BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="CaseNo" DataSourceID="sdsViolationCaseP_99"
            GridLines="None" AllowPaging="True" PageSize="5" Width="100%">
            <Columns>
                <asp:BoundField DataField="CaseNo" HeaderText="CaseNo" ReadOnly="True" SortExpression="CaseNo" Visible="False" />
                <asp:BoundField DataField="CountByType" HeaderText="案件總數(依類別)" ReadOnly="True" SortExpression="CountByType" />
                <asp:BoundField DataField="CountByDep" HeaderText="各站總數(依類別)" ReadOnly="True" SortExpression="CountByDep" />
                <asp:BoundField DataField="CaseType" HeaderText="CaseType" SortExpression="CaseType" Visible="False" />
                <asp:BoundField DataField="CaseType_C" HeaderText="類別" ReadOnly="True" SortExpression="CaseType_C" />
                <asp:BoundField DataField="DepNo" HeaderText="站別編號" SortExpression="DepNo" />
                <asp:BoundField DataField="DepName" HeaderText="站別" SortExpression="DepName" />
                <asp:BoundField DataField="LinesNo" HeaderText="路線" SortExpression="LinesNo" />
                <asp:BoundField DataField="Driver" HeaderText="Driver" SortExpression="Driver" Visible="False" />
                <asp:BoundField DataField="DriverName" HeaderText="違規駕駛" SortExpression="DriverName" />
                <asp:BoundField DataField="ViolationDate" DataFormatString="{0:D}" HeaderText="違規日期" SortExpression="ViolationDate" />
                <asp:BoundField DataField="ViolationRule" HeaderText="違規法條" SortExpression="ViolationRule" />
                <asp:BoundField DataField="FineAmount" HeaderText="裁罰金額" SortExpression="FineAmount" />
                <asp:BoundField DataField="ViolationPoint" HeaderText="記(扣)點" SortExpression="ViolationPoint" />
                <asp:BoundField DataField="ViolationNote" HeaderText="違規事項說明" SortExpression="ViolationNote" />
                <asp:BoundField DataField="TicketTitle" HeaderText="裁罰主旨" SortExpression="TicketTitle" />
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
    <asp:SqlDataSource ID="sdsViolationCaseP_DepNo" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, CaseType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = v.CaseType) AND (FKEY = '違規記錄檔      ViolationCase   CaseType')) AS CaseType_C, DepNo, DepName, ViolationDate, Driver, DriverName, ViolationRule, ViolationPoint, ViolationNote FROM ViolationCase AS v WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsViolationCaseP_Driver" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, CaseType, DepNo, DepName, Car_ID, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = v.CaseType) AND (FKEY = '違規記錄檔      ViolationCase   CaseType')) AS CaseType_C, ViolationDate, Driver, DriverName, ViolationRule, ViolationPoint, ViolationNote FROM ViolationCase AS v WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsViolationCaseP_Point" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, CaseType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.CaseType) AND (FKEY = '違規記錄檔      ViolationCase   CaseType')) AS CaseType_C, DepNo, DepName, ViolationDate, DriverName, ViolationRule, ViolationPoint, ViolationNote, CASE WHEN DateDiff(month , ViolationDate , '2019/07/01') &lt;= 6 THEN 'V' ELSE '' END AS IsAffectivePoint FROM ViolationCase AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsViolationCaseP_99" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, (SELECT COUNT(CaseNo) AS Expr1 FROM ViolationCase WHERE (CaseType = a.CaseType)) AS CountByType, (SELECT COUNT(CaseNo) AS Expr1 FROM ViolationCase AS ViolationCase_1 WHERE (DepNo = a.DepNo)) AS CountByDep, CaseType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.CaseType) AND (FKEY = '違規記錄檔      ViolationCase   CaseType')) AS CaseType_C, DepNo, DepName, LinesNo, Driver, DriverName, ViolationDate, ViolationRule, FineAmount, ViolationPoint, ViolationNote, TicketTitle FROM ViolationCase AS a WHERE (1 &lt;&gt; 1) ORDER BY DepNo, ViolationDate"></asp:SqlDataSource>
</asp:Content>
