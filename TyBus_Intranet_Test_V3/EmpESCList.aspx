<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmpESCList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmpESCList" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="EmpESCListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">員工請假狀況月報</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel" Width="100%">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbYearMonth_Search" runat="server" CssClass="text-Right-Blue" Text="查詢年月：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eYearMonth_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Label ID="lbNoteText_Search" runat="server" CssClass="text-Left-Red" Text="請輸入當月任一日期" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDepartment_Search" runat="server" CssClass="text-Right-Blue" Text="查詢單位：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" OnTextChanged="eDepNo_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eDepName_Search" runat="server" CssClass="text-Left-Black" Width="60%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbOK" runat="server" CssClass="button-Blue" Text="查詢" OnClick="bbOK_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Black" Text="列印" OnClick="bbPrint_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="離開" OnClick="bbClose_Click" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="plShow" runat="server" CssClass="ShowPanel">
        <asp:Calendar ID="calShowdata" runat="server" BackColor="White" BorderColor="White" BorderWidth="1px" Font-Names="Verdana" Font-Size="9pt" ForeColor="Black" Width="100%" NextPrevFormat="FullMonth" SelectionMode="None" OnDayRender="calShowdata_DayRender">
            <DayHeaderStyle Font-Bold="True" Font-Size="8pt" />
            <NextPrevStyle Font-Size="8pt" ForeColor="#333333" Font-Bold="True" VerticalAlign="Bottom" />
            <OtherMonthDayStyle ForeColor="#999999" />
            <SelectedDayStyle BackColor="#333399" ForeColor="White" />
            <TitleStyle BackColor="White" BorderColor="Black" BorderWidth="4px" Font-Bold="True" Font-Size="12pt" ForeColor="#333399" />
            <TodayDayStyle BackColor="#CCCCCC" />
        </asp:Calendar>
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
            <LocalReport ReportPath="Report\AnecdoteCaseP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
