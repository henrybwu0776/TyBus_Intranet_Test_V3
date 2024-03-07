<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmpWorkList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmpWorkList" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="EmpWorkListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">票價補助發放名冊(內勤)</a>
    </div>
    <br />
    <asp:Panel ID="plCalculater" runat="server" CssClass="ShowPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:Label ID="lbCalYM" runat="server" CssClass="text-Right-Blue" Text="計算年月(西元年 4 碼+月份 2 碼)" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCalYM" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalBase" runat="server" CssClass="text-Right-Blue" Text="內勤金額" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCalBase" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalBase_Driver" runat="server" CssClass="text-Right-Blue" Text="駕駛金額" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCalBase_Driver" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col" colspan="4" />
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbCal" runat="server" CssClass="button-Blue" Text="計算" Width="95%" OnClick="bbCal_Click" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出EXCEL" Width="95%" OnClick="bbExcel_Click" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Black" Text="預覽" Width="95%" OnClick="bbPrint_Click" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" Width="95%" OnClick="bbExit_Click" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="plShow" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridResultList" runat="server" Width="100%"></asp:GridView>
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
            <LocalReport ReportPath="Report\EmpWorkListP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
