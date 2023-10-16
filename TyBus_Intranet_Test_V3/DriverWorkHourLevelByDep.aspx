<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DriverWorkHourLevelByDep.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DriverWorkHourLevelByDep" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="DriverWorkHourLevelByDepForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">駕駛員工時級數分析表(單位)</a>
    </div>
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDriveDate_Search" runat="server" CssClass="text-Right-Blue" Text="行車日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDriveYear_Search" runat="server" CssClass="text-Left-Black" Width="50%" />
                    <asp:DropDownList ID="eDriveMonth_Search" runat="server" CssClass="text-Left-Black" Width="30%">
                        <asp:ListItem Text="1" Value="01" />
                        <asp:ListItem Text="2" Value="02" />
                        <asp:ListItem Text="3" Value="03" />
                        <asp:ListItem Text="4" Value="04" />
                        <asp:ListItem Text="5" Value="05" />
                        <asp:ListItem Text="6" Value="06" />
                        <asp:ListItem Text="7" Value="07" />
                        <asp:ListItem Text="8" Value="08" />
                        <asp:ListItem Text="9" Value="09" />
                        <asp:ListItem Text="10" Value="10" />
                        <asp:ListItem Text="11" Value="11" />
                        <asp:ListItem Text="12" Value="12" />
                    </asp:DropDownList>
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:DropDownList ID="eDepNoS_Search" runat="server" CssClass="text-Left-Black" Width="40%" DataSourceID="sdsDepNoList_S" DataTextField="DepName" DataValueField="DEPNO" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="titleText-S-Black" Text="～" />
                    <asp:DropDownList ID="eDepNoE_Search" runat="server" CssClass="text-Left-Black" Width="40%" DataSourceID="sdsDepNoList_E" DataTextField="DepName" DataValueField="DEPNO" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col" colspan="5" />
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Red" OnClick="bbPrint_Click" Text="列印" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" OnClick="bbExcel_Click" Text="匯出 EXCEL" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" OnClick="bbClose_Click" Text="結束" Width="95%" />
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
            <LocalReport ReportPath="Report\DriverWorkHourLevelByDepP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsDepNoList_S" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CAST('' AS varchar) AS DepNo, CAST('' AS varchar) AS DepName UNION ALL SELECT DEPNO, DEPNO + '_' + NAME AS DepName FROM DEPARTMENT WHERE (ISNULL(InSHReport, 'X') = 'V')"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsDepNoList_E" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CAST('' AS varchar) AS DepNo, CAST('' AS varchar) AS DepName UNION ALL SELECT DEPNO, DEPNO + '_' + NAME AS DepName FROM DEPARTMENT WHERE (ISNULL(InSHReport, 'X') = 'V')"></asp:SqlDataSource>
</asp:Content>
