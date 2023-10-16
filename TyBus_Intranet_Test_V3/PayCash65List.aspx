<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="PayCash65List.aspx.cs" Inherits="TyBus_Intranet_Test_V3.PayCash65List" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="PayCash65ListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">精勤保安獎金減發名冊</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalYM_Search" runat="server" CssClass="text-Right-Blue" Text="計算年月：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eCalYear_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbCalYear_Search" runat="server" CssClass="text-Left-Black" Text=" 年" Width="10%" />
                    <asp:TextBox ID="eCalMonth_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbCalMonth_Search" runat="server" CssClass="text-Left-Black" Text=" 月" Width="10%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Black" Text="列印" OnClick="bbPrint_Click" Width="95%" />
                </td>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="95%" />
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
        <asp:GridView ID="gridPayCash65List" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataSourceID="sdsPayCash65List" GridLines="None" Width="100%">
            <Columns>
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" ReadOnly="True" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepNo_L" HeaderText="部門" ReadOnly="True" SortExpression="DepNo_L" />
                <asp:BoundField DataField="driver" HeaderText="駕駛員工號" ReadOnly="True" SortExpression="driver" />
                <asp:BoundField DataField="EmpName" HeaderText="姓名" ReadOnly="True" SortExpression="EmpName" />
                <asp:BoundField DataField="EmpTitle" HeaderText="職稱" ReadOnly="True" SortExpression="EmpTitle" />
                <asp:BoundField DataField="car_no" HeaderText="車輛代碼" ReadOnly="True" SortExpression="car_no" />
                <asp:BoundField DataField="car_id" HeaderText="車牌號碼" ReadOnly="True" SortExpression="car_id" />
                <asp:BoundField DataField="HapDate" DataFormatString="{0:d}" HeaderText="發生日期" ReadOnly="True" SortExpression="HapDate" />
                <asp:BoundField DataField="Content" HeaderText="減發內容" ReadOnly="True" SortExpression="Content" />
                <asp:BoundField DataField="AMT_Duty" HeaderText="行車獎金扣額" ReadOnly="True" SortExpression="AMT_Duty" />
                <asp:BoundField DataField="AMT_FixAsk" HeaderText="車損扣額" ReadOnly="True" SortExpression="AMT_FixAsk" />
                <asp:BoundField DataField="Pay_Day" DataFormatString="{0:d}" HeaderText="減發日期" ReadOnly="True" SortExpression="Pay_Day" />
                <asp:BoundField DataField="WorkType" HeaderText="任職情況" ReadOnly="True" SortExpression="WorkType" />
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
            <LocalReport ReportPath="Report\PayCash65ListP.rdlc">
                <DataSources>
                </DataSources>
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsPayCash65List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT DepNo, (SELECT NAME FROM DEPARTMENT AS z WHERE (DEPNO = aa.DepNo)) AS DepNo_L, driver, EmpName, EmpTitle, car_no, car_id, HapDate, Content, AMT_Duty, AMT_FixAsk, Pay_Day, (SELECT CAST(worktype AS varchar(4)) AS WorkType FROM EMPLOYEE AS e5 WHERE (EMPNO = aa.driver)) AS WorkType FROM (SELECT DepNo, driver, (SELECT NAME FROM EMPLOYEE AS e1 WHERE (EMPNO = a.driver)) AS EmpName, (SELECT (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '人事資料檔      EMPLOYEE        TITLE') AND (CLASSNO = e2.TITLE)) AS emptitle FROM EMPLOYEE AS e2 WHERE (EMPNO = a.driver)) AS EmpTitle, car_no, car_id, day1 AS HapDate, (SELECT [Content] FROM DutyCode AS dc WHERE (DutyCode = a.DutyCode)) AS Content, amt AS AMT_Duty, CAST(0 AS float) AS AMT_FixAsk, day2 AS Pay_Day FROM Duty AS a WHERE (1 &lt;&gt; 1) UNION ALL SELECT c.DEPNO, b.Empno AS Driver, (SELECT NAME FROM EMPLOYEE AS e3 WHERE (EMPNO = b.Empno)) AS EmpName, (SELECT (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (FKEY = '人事資料檔      EMPLOYEE        TITLE') AND (CLASSNO = e4.TITLE)) AS emptitle FROM EMPLOYEE AS e4 WHERE (EMPNO = b.Empno)) AS EmpTitle, c.Car_no, c.CAR_ID, c.DAMDATE AS HapDate, c.DAMPART AS Content, CAST(0 AS float) AS AMT_Duty, ISNULL(b.Amt, 0) AS AMT_FixAsk, b.PayDate FROM FixaskB AS b LEFT OUTER JOIN FIXASK AS c ON c.FIXASKNO = b.FixaskNo WHERE (1 &lt;&gt; 1)) AS aa ORDER BY DepNo, driver"></asp:SqlDataSource>
</asp:Content>
