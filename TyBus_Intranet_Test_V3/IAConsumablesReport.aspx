<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="IAConsumablesReport.aspx.cs" Inherits="TyBus_Intranet_Test_V3.IAConsumablesReport" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="IAConsumablesReportForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">電腦課耗材報表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColBorder ColHeight ColWidth-5Col">
                    <asp:Label ID="lbReportName" runat="server" CssClass="text-Right-Blue" Text="報表：" Width="95%" />
                </td>
                <td class="ColBorder ColHeight ColWidth-5Col">
                    <asp:DropDownList ID="ddlReportName" runat="server" CssClass="text-Left-Black" Width="95%" OnSelectedIndexChanged="ddlReportName_SelectedIndexChanged" AutoPostBack="True">
                        <asp:ListItem Value="P01" Text="電腦課耗材庫存清單" />
                        <asp:ListItem Value="P02" Text="耗材進貨統計表" />
                        <asp:ListItem Value="P03" Text="耗材領出統計表" />
                        <asp:ListItem Value="P04" Text="申購統計表" />
                    </asp:DropDownList>
                </td>
                <td class="ColBorder ColHeight ColWidth-5Col">
                    <asp:Label ID="lbReportDate" runat="server" CssClass="text-Right-Blue" Text="日期區間：" Width="95%" />
                </td>
                <td class="ColBorder ColHeight ColWidth-5Col" colspan="2">
                    <asp:TextBox ID="eStartDate_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eEndDate_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="lbConsNo_Search" runat="server" CssClass="text-Right-Blue" Text="耗材編號：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                    <asp:TextBox ID="eConsNo_S_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eConsNo_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="lbBrand_Search" runat="server" CssClass="text-Right-Blue" Text="廠牌：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:DropDownList ID="eBrand_Search" runat="server" CssClass="text-Left-Black" Width="95%" DataSourceID="dsBrand" DataTextField="ClassTxt" DataValueField="ClassNo" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="lbSystematics_Search" runat="server" CssClass="text-Right-Blue" Text="品項：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:DropDownList ID="eSystematics_Search" runat="server" CssClass="text-Left-Black" Width="95%" DataSourceID="dsSystematics" DataTextField="ClassTxt" DataValueField="ClassNo" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="lbConsName_Search" runat="server" CssClass="text-Right-Blue" Text="品名：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                    <asp:TextBox ID="eConsName_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="lbCorrespondModel_Search" runat="server" CssClass="text-Right-Blue" Text="適用機型：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:TextBox ID="eCorrespondModel_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="lbConsStatus_Search" runat="server" CssClass="text-Right-Blue" Text="庫存狀況：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:DropDownList ID="eConsStatus_Search" runat="server" CssClass="text-Left-Black" Width="95%">
                        <asp:ListItem Value="S00" Text="" Selected="True" />
                        <asp:ListItem Value="S01" Text="高於安全量" />
                        <asp:ListItem Value="S02" Text="低於安全量" />
                        <asp:ListItem Value="S03" Text="已無庫存" />
                        <asp:ListItem Value="S05" Text="採購中" />
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-5Col" colspan="2" />
                <td class="ColHeight ColWidth-5Col">
                    <asp:Button ID="bbPreview" runat="server" CssClass="button-Black" Text="預覽" OnClick="bbPreview_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-5Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-5Col">
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" OnClick="bbExit_Click" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-5Col" />
                <td class="ColWidth-5Col" />
                <td class="ColWidth-5Col" />
                <td class="ColWidth-5Col" />
                <td class="ColWidth-5Col" />
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
            <LocalReport ReportPath="Report\IAConsumablesP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <!--廠商代碼對照資料庫-->
    <asp:SqlDataSource ID="dsBrand" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select null as ClassNo, null as ClassTxt
union all
select ClassNo, ClassTxt 
  from DBDICB
 where FKey = '電腦課耗材管理  fmIAConsumables Brand'"></asp:SqlDataSource>
    <!--品項代碼對照資料庫-->
    <asp:SqlDataSource ID="dsSystematics" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select null as ClassNo, null as ClassTxt
union all
select ClassNo, ClassTxt 
  from DBDICB
 where FKey = '電腦課耗材管理  fmIAConsumables Systematics'"></asp:SqlDataSource>
</asp:Content>
