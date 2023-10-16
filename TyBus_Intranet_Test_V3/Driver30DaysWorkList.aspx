<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="Driver30DaysWorkList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.Driver30DaysWorkList" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="Driver30DaysWorkListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">駕駛員最近30天出勤狀況</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:DropDownList ID="ddlDepNo" runat="server" CssClass="text-Left-Black" OnSelectedIndexChanged="ddlDepNo_SelectedIndexChanged" Width="90%" AutoPostBack="True"
                        DataSourceID="sdsDepNoList" DataTextField="DepName" DataValueField="DepNo" Height="16px">
                    </asp:DropDownList>
                    <asp:TextBox ID="eDepNo" runat="server" CssClass="text-Left-Black" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDriver" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eDriver" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_TextChanged" Width="30%" />
                    <asp:Label ID="lbDriverName" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbCalDays" runat="server" CssClass="text-Right-Blue" Text="回算天數：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eCalDays" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbLastDays" runat="server" CssClass="text-Right-Blue" Text="最近天數：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eLastDays" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbStopDate" runat="server" CssClass="text-Right-Blue" Text="截止日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eStopDate" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbPrint_OverTime" runat="server" CssClass="button-Red" Text="超時人員列表" OnClick="bbPrint_OverTime_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbExcel_OverTime" runat="server" CssClass="button-Blue" Text="匯出超時人員" OnClick="bbExcel_OverTime_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
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
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:Label ID="lbDataNote" runat="server" CssClass="errorMessageText" Text="" />
        <asp:GridView ID="gridShowData" runat="server" BackColor="White" BorderColor="Black" BorderStyle="Double" BorderWidth="1px"
            CellPadding="4" CssClass="text-Left-Black" Width="100%" OnRowDataBound="gridShowData_RowDataBound">
            <EditRowStyle BorderStyle="None" />
            <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
            <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
            <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
            <RowStyle BackColor="White" ForeColor="Black" BorderColor="Black" BorderStyle="Double" BorderWidth="1px" />
            <SelectedRowStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
            <SortedAscendingCellStyle BackColor="#EDF6F6" />
            <SortedAscendingHeaderStyle BackColor="#0D4AC4" />
            <SortedDescendingCellStyle BackColor="#D6DFDF" />
            <SortedDescendingHeaderStyle BackColor="#002876" />
        </asp:GridView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsDepNoList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DepNo)) AS DepName FROM EmployeeDepNo AS e WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
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
            <LocalReport ReportPath="Report\DriverWorkOver7DaysP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
