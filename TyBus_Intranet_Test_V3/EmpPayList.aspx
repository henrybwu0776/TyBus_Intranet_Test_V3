<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmpPayList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmpPayList" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="EmpPayListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">到離員工應發薪資清冊</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbPayYM_Search" runat="server" CssClass="text-Right-Blue" Text="發薪年月：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="ePayYear_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbPayYear_Search" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:DropDownList ID="ePayMonth_Search" runat="server" CssClass="text-Left-Black" Width="30%">
                        <asp:ListItem Text="1" Value="1" />
                        <asp:ListItem Text="2" Value="2" />
                        <asp:ListItem Text="3" Value="3" />
                        <asp:ListItem Text="4" Value="4" />
                        <asp:ListItem Text="5" Value="5" />
                        <asp:ListItem Text="6" Value="6" />
                        <asp:ListItem Text="7" Value="7" />
                        <asp:ListItem Text="8" Value="8" />
                        <asp:ListItem Text="9" Value="9" />
                        <asp:ListItem Text="10" Value="10" />
                        <asp:ListItem Text="11" Value="11" />
                        <asp:ListItem Text="12" Value="12" />
                    </asp:DropDownList>
                    <asp:Label ID="lbPayMonth_Search" runat="server" CssClass="text-Left-Black" Text=" 月" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbPayDur_Search" runat="server" CssClass="text-Right-Blue" Text="期別：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ePayDur_Search" runat="server" CssClass="text-Left-Black" Width="95%">
                        <asp:ListItem Selected="True" Text="1" Value="1" />
                        <asp:ListItem Text="2" Value="2" />
                        <asp:ListItem Text="3" Value="3" />
                    </asp:DropDownList>
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" OnClick="bbSearch_Click" Text="查詢" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Blue" OnClick="bbPrint_Click" Text="列印" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" OnClick="bbExcel_Click" Text="匯出EXCEL" Width="95%" />
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
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridEmpPayList" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataSourceID="sdsEmpPayList" GridLines="None" Width="100%">
            <Columns>
                <asp:BoundField DataField="depno" HeaderText="depno" SortExpression="depno" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="單位" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="empno" HeaderText="工號" SortExpression="empno" />
                <asp:BoundField DataField="name" HeaderText="姓名" SortExpression="name" />
                <asp:BoundField DataField="worktype" HeaderText="狀態" SortExpression="worktype" />
                <asp:BoundField DataField="AssumeDay" DataFormatString="{0:D}" HeaderText="到職日" ReadOnly="True" SortExpression="AssumeDay" />
                <asp:BoundField DataField="LEAVEDAY" DataFormatString="{0:D}" HeaderText="離職日" SortExpression="LEAVEDAY" />
                <asp:BoundField DataField="BEGINDAY" DataFormatString="{0:D}" HeaderText="留停起日" ReadOnly="True" SortExpression="BEGINDAY" />
                <asp:BoundField DataField="STOPDAY" DataFormatString="{0:D}" HeaderText="留停迄日" SortExpression="STOPDAY" />
                <asp:BoundField DataField="WorkDays" HeaderText="上班天數" ReadOnly="True" SortExpression="WorkDays" />
                <asp:BoundField DataField="givcash" HeaderText="應發金額" SortExpression="givcash" />
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
            <LocalReport ReportPath="Report\EmpPayListP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsEmpPayList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT p.depno, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = p.depno)) AS DepName, p.empno, p.name, e.worktype, CASE WHEN e.AssumeDay &gt;= '2020/09/01' THEN AssumeDay ELSE NULL END AS AssumeDay, e.LEAVEDAY, e.BEGINDAY, e.STOPDAY, DATEDIFF(Day, CASE WHEN e.AssumeDay &gt;= '2020/09/01' THEN AssumeDay ELSE '2020/09/01' END, CASE WHEN isnull(e.LeaveDay , '2020/09/30') &gt; '2020/09/30' THEN '2020/09/30' ELSE isnull(e.LeaveDay , '2020/09/30') END) + 1 AS WorkDays, p.givcash FROM PAYREC AS p LEFT OUTER JOIN EMPLOYEE AS e ON e.EMPNO = p.empno WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
