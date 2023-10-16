<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CustomServiceReportList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CustomServiceReportList" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="CustomServiceReportListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">0800客服專線統計報表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                    <asp:RadioButtonList ID="rbGroupBy" runat="server" CssClass="text-Left-Black" AutoPostBack="true" RepeatColumns="4" Width="90%">
                        <asp:ListItem Text="依單位別" Value="byDepNo" Selected="True" />
                        <asp:ListItem Text="依駕駛別" Value="byDriver" />
                        <asp:ListItem Text="依反映事項" Value="byServiceType" />
                        <asp:ListItem Text="依路線別" Value="byLinesNo" />
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalDate" runat="server" CssClass="text-Right-Blue" Text="開單日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eCalDate_Start" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eCalDate_End" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbServiceType" runat="server" CssClass="text-Right-Blue" Text="反映事項：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlServiceType" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True"
                        DataSourceID="sdsServiceType" DataTextField="TypeText" DataValueField="TypeLevel1">
                    </asp:DropDownList>
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:RadioButtonList ID="rbSelectType" runat="server" CssClass="text-Left-Black" AutoPostBack="true" RepeatColumns="3" Width="90%">
                        <asp:ListItem Text="只顯示未結案" Value="0" />
                        <asp:ListItem Text="只顯示已結案" Value="1" />
                        <asp:ListItem Text="全部資料" Value="2" Selected="True" />
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbPreview" runat="server" CssClass="button-Black" Text="預覽報表" OnClick="bbPreview_Click" Width="90%" />
                </td>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
                </td>
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plReport" runat="server" CssClass="PrintPanel">
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
            <LocalReport ReportPath="Report\ServiceCountByDapNo.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridCustomServiceP" runat="server" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" GridLines="None">
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
    <asp:SqlDataSource ID="sdsCustomServiceP" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT 單位編號, 單位, 反映事項代號, 反映事項, CASE WHEN isnull(客訴事項代號 , '') = '' THEN 反映事項代號 ELSE 客訴事項代號 END AS 客訴事項代號, CASE WHEN isnull(客訴事項 , '') = '' THEN 反映事項 ELSE 客訴事項 END AS 客訴事項, CASE WHEN (isnull(客訴類別代號 , '') = '') AND (isnull(客訴事項代號 , '') = '') THEN 反映事項代號 WHEN (isnull(客訴類別代號 , '') = '') AND (isnull(客訴事項代號 , '') &lt;&gt; '') THEN 客訴事項代號 ELSE 客訴類別代號 END AS 客訴類別代號, CASE WHEN (isnull(客訴類別 , '') = '') AND (isnull(客訴事項 , '') = '') THEN 反映事項 WHEN (isnull(客訴類別 , '') = '') AND (isnull(客訴事項 , '') &lt;&gt; '') THEN 客訴事項 ELSE 客訴類別 END AS 客訴類別, 路線代號, 駕駛員代號, 駕駛員 FROM (SELECT ISNULL(AthorityDepNo, '') AS 單位編號, ISNULL(AthorityDepNote, '無單位') AS 單位, ISNULL(ServiceType, '') AS 反映事項代號, (SELECT TypeText FROM CustomServiceType WHERE (TypeStep = '1') AND (TypeLevel1 = a.ServiceType)) AS 反映事項, ISNULL(ServiceTypeB, '') AS 客訴事項代號, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE (TypeStep = '2') AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB)) AS 客訴事項, ISNULL(ServiceTypeC, '') AS 客訴類別代號, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE (TypeStep = '3') AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB) AND (TypeLevel3 = a.ServiceTypeC)) AS 客訴類別, ISNULL(LinesNo, '無路線別') AS 路線代號, ISNULL(Driver, '') AS 駕駛員代號, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.Driver)) AS 駕駛員 FROM CustomService AS a WHERE (1 = 1) UNION ALL SELECT ISNULL(AthorityDepNo2, '') AS 單位編號, ISNULL(AthorityDepNote2, '無單位') AS 單位, ISNULL(ServiceType, '') AS 反映事項代號, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_3 WHERE (TypeStep = '1') AND (TypeLevel1 = a.ServiceType)) AS 反映事項, ISNULL(ServiceTypeB, '') AS 客訴事項代號, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE (TypeStep = '2') AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB)) AS 客訴事項, ISNULL(ServiceTypeC, '') AS 客訴類別代號, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE (TypeStep = '3') AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB) AND (TypeLevel3 = a.ServiceTypeC)) AS 客訴類別, ISNULL(LinesNo2, '無路線別') AS 路線代號, ISNULL(Driver2, '') AS 駕駛員代號, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.Driver2)) AS 駕駛員 FROM CustomService AS a WHERE (ISNULL(AthorityDepNo2, N'') &lt;&gt; '')) AS f ORDER BY 單位編號, 反映事項代號, 客訴事項代號, 客訴類別代號, 路線代號, 駕駛員代號"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsServiceType" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST('' AS nvarchar) AS TypeLevel1, CAST('' AS nvarchar) AS TypeText UNION ALL SELECT TypeLevel1, TypeText FROM CustomServiceType WHERE (TypeStep = '1') ORDER BY TypeLevel1"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsCountDepNo" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT AthorityDepNo, AthorityDepNo_C, COUNT(AthorityDepNo) AS RCount FROM (SELECT ISNULL(AthorityDepNo, 'ZZ') AS AthorityDepNo, CASE WHEN AthorityDepNo = '99' THEN '其他' ELSE ISNULL(AthorityDepNote , '不分站別') END AS AthorityDepNo_C FROM CustomService AS a UNION ALL SELECT ISNULL(AthorityDepNo2, 'ZZ') AS AthorityDepNo, CASE WHEN AthorityDepNo2 = '99' THEN '其他' ELSE ISNULL(AthorityDepNote2 , '不分站別') END AS AthorityDepNo_C FROM CustomService AS a WHERE (ISNULL(AthorityDepNo2, N'') &lt;&gt; '')) AS z GROUP BY AthorityDepNo, AthorityDepNo_C ORDER BY AthorityDepNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsCountServiceType" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT ServiceType, (SELECT TypeText FROM CustomServiceType WHERE (TypeStep = '1') AND (TypeLevel1 = CustomService.ServiceType)) AS ServiceType_C, COUNT(ServiceNo) AS RCount FROM CustomService GROUP BY ServiceType ORDER BY ServiceType"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsCountServiceTypeB" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT ServiceType, ServiceType_C, COUNT(ServiceNo) AS RCount FROM (SELECT ServiceNo, ServiceType + ISNULL(ServiceTypeB, '000') AS ServiceType, CASE WHEN ISNULL(ServiceTypeB , '') &lt;&gt; '' THEN (SELECT TypeText FROM CustomServiceType WHERE TypeStep = '2' AND TypeLevel1 = a.ServiceType AND TypeLevel2 = a.ServiceTypeB) ELSE (SELECT TypeText FROM CustomserviceType WHERE TypeStep = '1' AND TypeLevel1 = a.ServiceType) END AS ServiceType_C FROM CustomService AS a) AS z GROUP BY ServiceType, ServiceType_C ORDER BY ServiceType, ServiceType_C"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsCountServiceTypeC" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT ServiceType, ServiceType_C, COUNT(ServiceNo) AS RCount FROM (SELECT ServiceNo, ServiceType + ISNULL(ServiceTypeB, '000') + ISNULL(ServiceTypeC, '000') AS ServiceType, (SELECT TypeText FROM CustomServiceType WHERE (TypeNo = a.ServiceType + ISNULL(a.ServiceTypeB, N'000') + ISNULL(a.ServiceTypeC, N'000'))) AS ServiceType_C FROM CustomService AS a) AS z GROUP BY ServiceType, ServiceType_C ORDER BY ServiceType, ServiceType_C"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsCountDriver" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT Driver, Driver_C, COUNT(Driver) AS RCount FROM (SELECT ISNULL(Driver, '000000') AS Driver, ISNULL((SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.Driver)), '不分駕駛員') AS Driver_C FROM CustomService AS a UNION ALL SELECT ISNULL(Driver2, '000000') AS Driver, ISNULL((SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.Driver2)), '不分駕駛員') AS Driver_C FROM CustomService AS a WHERE (ISNULL(Driver2, N'') &lt;&gt; '')) AS z GROUP BY Driver, Driver_C ORDER BY Driver"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsCountLinesNo" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT LinesNo, COUNT(LinesNo) AS RCount FROM (SELECT ISNULL(LinesNo, '不分路線') AS LinesNo FROM CustomService AS a WHERE (1 = 1) UNION ALL SELECT LinesNo2 AS LinesNo FROM CustomService AS a WHERE (ISNULL(LinesNo2, N'') &lt;&gt; '')) AS z GROUP BY LinesNo ORDER BY LinesNo"></asp:SqlDataSource>
</asp:Content>
