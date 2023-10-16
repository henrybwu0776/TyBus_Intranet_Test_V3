<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmpMeritsPrint.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmpMeritsPrint" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="EmpMeritsPrintForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">人事考績獎懲明細表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbListYM_Search" runat="server" CssClass="text-Right-Blue" Text="考核年月" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eListYear_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbListYear_Search" runat="server" CssClass="text-Left-Blue" Text=" 年 " />
                    <asp:DropDownList ID="ddlListMonth_Search" runat="server" CssClass="text-Left-Black" Width="35%">
                        <asp:ListItem Value="01" Text="一月" />
                        <asp:ListItem Value="02" Text="二月" />
                        <asp:ListItem Value="03" Text="三月" />
                        <asp:ListItem Value="04" Text="四月" />
                        <asp:ListItem Value="05" Text="五月" />
                        <asp:ListItem Value="06" Text="六月" />
                        <asp:ListItem Value="07" Text="七月" />
                        <asp:ListItem Value="08" Text="八月" />
                        <asp:ListItem Value="09" Text="九月" />
                        <asp:ListItem Value="10" Text="十月" />
                        <asp:ListItem Value="11" Text="十一月" />
                        <asp:ListItem Value="12" Text="十二月" />
                    </asp:DropDownList>
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbFNo_Search" runat="server" CssClass="text-Right-Blue" Text="字號" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eFNo_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbFDate_Search" runat="server" CssClass="text-Right-Blue" Text="發文日期" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eFDate_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" colspan="2">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="120px" />
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Blue" Text="列印" OnClick="bbPrint_Click" Width="120px" />
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="離開" OnClick="bbClose_Click" Width="120px" />
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
    <asp:Panel ID="plShow" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridMeritsList" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px" CellPadding="4" DataSourceID="dsMeritsList" Width="100%">
            <Columns>
                <asp:BoundField DataField="EmpNo" HeaderText="員工編號" ReadOnly="True" SortExpression="EmpNo" />
                <asp:BoundField DataField="Title_C" HeaderText="職稱" ReadOnly="True" SortExpression="Title_C" />
                <asp:BoundField DataField="Name" HeaderText="姓名" ReadOnly="True" SortExpression="Name" />
                <asp:BoundField DataField="BDate" HeaderText="年/月/日" ReadOnly="True" SortExpression="BDate" DataFormatString="{0:D}" />
                <asp:BoundField DataField="BFact1" HeaderText="事實陳述" ReadOnly="True" SortExpression="BFact1" />
                <asp:BoundField DataField="Merit1" HeaderText="大功" ReadOnly="True" SortExpression="Merit1" />
                <asp:BoundField DataField="Merit2" HeaderText="小功" ReadOnly="True" SortExpression="Merit2" />
                <asp:BoundField DataField="Merit3" HeaderText="嘉獎" ReadOnly="True" SortExpression="Merit3" />
                <asp:BoundField DataField="Demerit1" HeaderText="大過" ReadOnly="True" SortExpression="Demerit1" />
                <asp:BoundField DataField="Demerit2" HeaderText="小過" ReadOnly="True" SortExpression="Demerit2" />
                <asp:BoundField DataField="Demerit3" HeaderText="申誡" ReadOnly="True" SortExpression="Demerit3" />
                <asp:BoundField DataField="Warning" HeaderText="警告" ReadOnly="True" SortExpression="Warning" />
                <asp:BoundField DataField="Alart" HeaderText="口頭警告" ReadOnly="True" SortExpression="Alart" />
                <asp:BoundField DataField="WRIWarning" HeaderText="書面警告" ReadOnly="True" SortExpression="WRIWarning" />
                <asp:BoundField DataField="Fcomment" HeaderText="其他決議" ReadOnly="True" SortExpression="Fcomment" />
            </Columns>
            <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
            <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
            <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
            <RowStyle BackColor="White" ForeColor="#003399" />
            <SelectedRowStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
            <SortedAscendingCellStyle BackColor="#EDF6F6" />
            <SortedAscendingHeaderStyle BackColor="#0D4AC4" />
            <SortedDescendingCellStyle BackColor="#D6DFDF" />
            <SortedDescendingHeaderStyle BackColor="#002876" />
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="plReport" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCloseReport" runat="server" CssClass="button-Red" OnClick="bbCloseReport_Click" Text="結束預覽" Width="120px" />
        <rsweb:ReportViewer ID="rvPrint" runat="server" Width="100%" BackColor="" ClientIDMode="AutoID" HighlightBackgroundColor="" InternalBorderColor="204, 204, 204" InternalBorderStyle="Solid" InternalBorderWidth="1px" LinkActiveColor="" LinkActiveHoverColor="" LinkDisabledColor="" PrimaryButtonBackgroundColor="" PrimaryButtonForegroundColor="" PrimaryButtonHoverBackgroundColor="" PrimaryButtonHoverForegroundColor="" SecondaryButtonBackgroundColor="" SecondaryButtonForegroundColor="" SecondaryButtonHoverBackgroundColor="" SecondaryButtonHoverForegroundColor="" SplitterBackColor="" ToolbarDividerColor="" ToolbarForegroundColor="" ToolbarForegroundDisabledColor="" ToolbarHoverBackgroundColor="" ToolbarHoverForegroundColor="" ToolBarItemBorderColor="" ToolBarItemBorderStyle="Solid" ToolBarItemBorderWidth="1px" ToolBarItemHoverBackColor="" ToolBarItemPressedBorderColor="51, 102, 153" ToolBarItemPressedBorderStyle="Solid" ToolBarItemPressedBorderWidth="1px" ToolBarItemPressedHoverBackColor="153, 187, 226">
            <LocalReport ReportPath="Report\EmpMeritsPrint.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="dsMeritsList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT CAST(NULL AS varchar) AS FNO, CAST(NULL AS datetime) AS FDATE, CAST(NULL AS varchar) AS DepNo, CAST(NULL AS varchar) AS DepNo_C, CAST(NULL AS varchar) AS EmpNo, CAST(NULL AS varchar) AS Title_C, CAST(NULL AS varchar) AS Name, CAST(NULL AS datetime) AS BDate, CAST(NULL AS varchar) AS BFact1, CAST(NULL AS float) AS Merit1, CAST(NULL AS float) AS Merit2, CAST(NULL AS float) AS Merit3, CAST(NULL AS float) AS Demerit1, CAST(NULL AS float) AS Demerit2, CAST(NULL AS float) AS Demerit3, CAST(NULL AS float) AS Warning, CAST(NULL AS float) AS Alart, CAST(NULL AS float) AS WRIWarning, CAST(NULL AS varchar) AS Fcomment"></asp:SqlDataSource>
</asp:Content>
