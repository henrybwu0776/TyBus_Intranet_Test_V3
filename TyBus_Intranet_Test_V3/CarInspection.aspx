<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CarInspection.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CarInspection" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="CarInspectionForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">車輛檢驗到期查核</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_S" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_E" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <br />
                    <asp:Label ID="eDepName_S" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="eDepName_E" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="4">
                    <asp:RadioButtonList ID="rbSearchMode" runat="server" CssClass="text-Left-Black" CellPadding="0" Width="90%" RepeatDirection="Horizontal">
                        <asp:ListItem Value="0" Text="所有車輛" Selected="True" />
                        <asp:ListItem Value="1" Text="過期尚未檢驗車輛" />
                        <asp:ListItem Value="2" Text="30天內需檢驗車輛" />
                    </asp:RadioButtonList>
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbCalCheckDate" runat="server" CssClass="button-Red" Text="重算檢驗日" OnClick="bbCalCheckDate_Click" UseSubmitBehavior="False" OnClientClick="this.disabled=true;" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Red" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Red" Text="列印報表" OnClick="bbPrint_Click" Width="90%" />
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
        <asp:GridView ID="gridCarInspectionList" runat="server" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataSourceID="sdsCarInspectionList" GridLines="None" AutoGenerateColumns="False" DataKeyNames="Car_No" Width="100%">
            <Columns>
                <asp:BoundField DataField="CompanyNo" HeaderText="CompanyNo" SortExpression="CompanyNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="所屬站別" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="Car_ID" HeaderText="牌照號碼" SortExpression="Car_ID" />
                <asp:BoundField DataField="Car_No" HeaderText="Car_No" ReadOnly="True" SortExpression="Car_No" Visible="False" />
                <asp:BoundField DataField="point" HeaderText="point" SortExpression="point" Visible="false" />
                <asp:BoundField DataField="Point_C" HeaderText="用途" SortExpression="Point_C" />
                <asp:BoundField DataField="ProdDate" DataFormatString="{0:d}" HeaderText="出廠日期" SortExpression="ProdDate" />
                <asp:BoundField DataField="StartDate" DataFormatString="{0:d}" HeaderText="可驗車日期" SortExpression="StartDate" />
                <asp:BoundField DataField="NCheckTerm" DataFormatString="{0:d}" HeaderText="檢驗期限" SortExpression="NCheckTerm" />
                <asp:BoundField DataField="nextedate" DataFormatString="{0:d}" HeaderText="下次檢驗日" SortExpression="nextedate" />
                <asp:BoundField DataField="LastCheckDate" DataFormatString="{0:d}" HeaderText="前次完成檢驗日" ReadOnly="True" SortExpression="LastCheckDate" />
                <asp:BoundField DataField="CheckDateDiff" HeaderText="上次驗車後間隔" ReadOnly="True" SortExpression="CheckDateDiff" DataFormatString="{0:N2}" />
                <asp:BoundField DataField="Car_Class_C" HeaderText="廠牌" ReadOnly="True" SortExpression="Car_Class_C" />
                <asp:BoundField DataField="Car_TypeID" HeaderText="型式" SortExpression="Car_TypeID" />
                <asp:BoundField DataField="getlicdate" DataFormatString="{0:d}" HeaderText="領照日期" SortExpression="getlicdate" />
                <asp:BoundField DataField="Class_C" HeaderText="類別" ReadOnly="True" SortExpression="Class_C" />
                <asp:BoundField DataField="Tran_Type_C" HeaderText="狀態" ReadOnly="True" SortExpression="Tran_Type_C" />
                <asp:BoundField DataField="Driver" HeaderText="Driver" SortExpression="Driver" Visible="False" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員" ReadOnly="True" SortExpression="DriverName" />
                <asp:BoundField DataField="driver1" HeaderText="driver1" SortExpression="driver1" Visible="False" />
                <asp:BoundField DataField="DriverName1" HeaderText="駕駛員二" ReadOnly="True" SortExpression="DriverName1" />
                <asp:BoundField DataField="ProdMonth" HeaderText="出廠月數" ReadOnly="True" SortExpression="ProdMonth" />
                <asp:BoundField DataField="UsedMonth" HeaderText="領照月數" ReadOnly="True" SortExpression="UsedMonth" />
                <asp:BoundField DataField="Intervals" HeaderText="驗車間隔" ReadOnly="True" SortExpression="Intervals" />
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
    <asp:SqlDataSource ID="sdsCarInspectionList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT ca.CompanyNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = ca.CompanyNo)) AS DepName, ca.Car_ID, ca.Car_No, ca.point, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = ca.point) AND (FKEY = '車輛資料作業    Car_infoA       POINT')) AS Point_C, ca.ProdDate, ca.NCheckTerm, ca.nextedate, cc.LastCheckDate, DATEADD(day, - 29, ca.nextedate) AS StartDate, DATEDIFF(month, ca.nextedate, cc.LastCheckDate) AS CheckDateDiff, (SELECT CLASSTXT FROM DBDICB AS DBDICB_3 WHERE (FKEY = '車輛資料作業    Car_infoA       CAR_CLASS') AND (CLASSNO = ca.Car_Class)) AS Car_Class_C, ca.Car_TypeID, ca.getlicdate, (SELECT CLASSTXT FROM DBDICB AS DBDICB_2 WHERE (FKEY = '車輛資料作業    Car_infoA       CLASS') AND (CLASSNO = ca.CLASS)) AS Class_C, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (FKEY = '車輛資料作業    Car_infoA       TRAN_TYPE') AND (CLASSNO = ca.Tran_Type)) AS Tran_Type_C, ca.Driver, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = ca.Driver)) AS DriverName, ca.driver1, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = ca.driver1)) AS DriverName1, DATEDIFF(month, ca.ProdDate, GETDATE()) AS ProdMonth, DATEDIFF(month, ca.getlicdate, GETDATE()) AS UsedMonth, CASE WHEN ca.Point IN ('X' , 'X2') THEN CASE WHEN DateDiff(month , ca.GetLicDate , GetDate()) &lt;= 60 THEN 60 WHEN DateDiff(month , ca.GetLicDate , GetDate()) &gt; 60 AND DateDiff(month , ca.GetLicDate , GetDate()) &lt;= 120 THEN 12 ELSE 6 END ELSE CASE WHEN DateDiff(month , ca.GetLicDate , GetDate()) &lt;= 60 THEN 12 WHEN DateDiff(month , ca.GetLicDate , GetDate()) &gt; 60 AND DateDiff(month , ca.GetLicDate , GetDate()) &lt;= 120 THEN 6 ELSE 4 END END AS Intervals FROM Car_infoA AS ca LEFT OUTER JOIN (SELECT MAX(CheckDate) AS LastCheckDate, Car_ID FROM Car_infoC GROUP BY Car_ID) AS cc ON cc.Car_ID = ca.Car_ID WHERE (1 &lt;&gt; 1) ORDER BY ca.CompanyNo, ca.Car_ID"></asp:SqlDataSource>

    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCLoseReport" runat="server" CssClass="button-Red" Text="關閉報表" OnClick="bbCLoseReport_Click" Width="120px" />
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
            <LocalReport ReportPath="Report\CarInspectionP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
