<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="AnnSignList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.AnnSignList" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
<asp:Content ID="AnnSignListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">公告簽名本</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbListTitle" runat="server" CssClass="text-Right-Blue" Text="公告標題" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eListTitle" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbBaseDate" runat="server" CssClass="text-Right-Blue" Text="基準日期" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eBaseDate" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:CheckBox ID="cbEmpType" runat="server" CssClass="text-Left-Blue" Text="包括駕駛" AutoPostBack="true" OnCheckedChanged="cbEmpType_CheckedChanged" />
                    <asp:Label ID="eEmpType" runat="server" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" colspan="4" />
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Blue" Text="列印出席簽到表" Width="120px" OnClick="bbPrint_Click" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" Width="120px" OnClick="bbExit_Click" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
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
            <LocalReport ReportPath="Report\AnnSignListP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
