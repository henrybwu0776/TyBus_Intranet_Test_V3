<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmpDutyHoursP.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmpDutyHoursP" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="EmpDutyHoursPForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">員工加班補休報表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" GroupingText="查詢條件" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbReportType" runat="server" CssClass="text-Right-Blue" Text="報表類別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="8">
                    <asp:RadioButtonList ID="rbReportType" runat="server" CssClass="text-Left-Black" RepeatColumns="4" AutoPostBack="true" Width="97%" OnSelectedIndexChanged="rbReportType_SelectedIndexChanged">
                        <asp:ListItem Text="加班時數統計表" Value="DutyHours" Selected="True" />
                        <asp:ListItem Text="換休時數統計表" Value="ESCHours" />
                        <asp:ListItem Text="逾期時數統計表" Value="Overdue" />
                        <asp:ListItem Text="可換休時數統計表" Value="UseableHours" />
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbCalYear" runat="server" CssClass="text-Right-Blue" Text="年度別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eCalYear_S" runat="server" CssClass="text-Left-Black" Width="15%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="年" />
                    <asp:TextBox ID="eCalMonth_S" runat="server" CssClass="text-Left-Black" Width="10%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="月 ～ " />
                    <asp:TextBox ID="eCalYear_E" runat="server" CssClass="text-Left-Black" Width="15%" />
                    <asp:Label ID="lbSplit_3" runat="server" CssClass="text-Left-Black" Text="年" />
                    <asp:TextBox ID="eCalMonth_E" runat="server" CssClass="text-Left-Black" Width="10%" />
                    <asp:Label ID="lbSplit_4" runat="server" CssClass="text-Left-Black" Text="月" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="單位別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eDepNo_S" runat="server" CssClass="text-Left-Black" Width="40%" OnTextChanged="eDepNo_S_TextChanged" AutoPostBack="true" />
                    <asp:Label ID="lbSplit_6" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_E" runat="server" CssClass="text-Left-Black" Width="40%" OnTextChanged="eDepNo_E_TextChanged" AutoPostBack="true" />
                    <br />
                    <asp:Label ID="eDepName_S" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="eDepName_E" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbEmpNo" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eEmpNo_S" runat="server" CssClass="text-Left-Black" Width="30%" OnTextChanged="eEmpNo_S_TextChanged" AutoPostBack="true" />
                    <asp:Label ID="eEmpName_S" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
            </tr>

            <tr>
                <td class="ColHeight ColWidth-9Col">
                    <asp:Button ID="bbPreview" runat="server" CssClass="button-Black" Text="預覽" OnClick="bbPreview_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-9Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出EXCEL" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-9Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
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
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
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
            <LocalReport ReportPath="Report\EmpDutyHoursP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
