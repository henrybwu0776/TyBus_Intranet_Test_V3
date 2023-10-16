<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="TimeTable.aspx.cs" Inherits="TyBus_Intranet_Test_V3.TimeTable" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="TimeTableForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">路線時刻表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbBuDate" runat="server" CssClass="text-Right-Blue" Text="日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eBuDate_S" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_3" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eBuDate_E" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class=",colh ColBorder ColWidth-9Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="單位：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eDepNo_S" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_E" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbLinesNo" runat="server" CssClass="text-Right-Blue" Text="路線：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eLinesNo_S" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eLinesNo_E" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-9Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-9Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出 Excel" OnClick="bbExcel_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-9Col">
                    <asp:Button ID="bbReport" runat="server" CssClass="button-Black" Text="預覽" OnClick="bbReport_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-9Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridTimeTable" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1"
            DataSourceID="sdsTimeTable" GridLines="None" Width="100%" PageSize="20">
            <Columns>
                <asp:BoundField DataField="BuDate" HeaderText="日期" ReadOnly="True" SortExpression="BuDate" DataFormatString="{0:d}" />
                <asp:BoundField DataField="DEPNO" HeaderText="DEPNO" SortExpression="DEPNO" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="單位" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="LinesNo" HeaderText="路線代號" SortExpression="LinesNo" />
                <asp:BoundField DataField="LinesName" HeaderText="路線名稱" ReadOnly="True" SortExpression="LinesName" />
                <asp:BoundField DataField="ToTime" HeaderText="出車時刻" SortExpression="ToTime" />
                <asp:BoundField DataField="ToLine" HeaderText="去程路線" SortExpression="ToLine" />
                <asp:BoundField DataField="BackTime" HeaderText="回程時刻" SortExpression="BackTime" />
                <asp:BoundField DataField="BackLine" HeaderText="回程路線" SortExpression="BackLine" />
                <asp:BoundField DataField="Car_Id" HeaderText="車號" SortExpression="Car_Id" />
                <asp:BoundField DataField="CarType_C" HeaderText="車種" SortExpression="CarType_C" />
                <asp:BoundField DataField="DRIVER" HeaderText="DRIVER" SortExpression="DRIVER" Visible="False" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛" SortExpression="DriverName" />
                <asp:BoundField DataField="CellPhone" HeaderText="電話" SortExpression="CellPhone" />
                <asp:BoundField DataField="ActualKM" HeaderText="公里數" SortExpression="ActualKM" />
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
            <LocalReport ReportPath="Report\TimeTableP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsTimeTable" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT a.BUDATE, a.DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DEPNO)) AS DepName, b.LinesNo, (SELECT LineName FROM Lines WHERE (LinesNo = b.LinesNo)) AS LinesName, b.ToTime, b.ToLine, b.BackTime, b.BackLine, b.Car_Id, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '行車記錄單b     runsheetb       CARTYPE') AND (CLASSNO = b.CarType)) AS CarType_C, a.DRIVER, e.NAME AS DriverName, e.CELLPHONE AS CellPhone, b.ActualKM FROM RUNSHEETA AS a LEFT OUTER JOIN RunSheetB AS b ON b.AssignNo = a.ASSIGNNO LEFT OUTER JOIN EMPLOYEE AS e ON e.EMPNO = a.DRIVER WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
