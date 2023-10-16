<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="WorkReportOverTimes.aspx.cs" Inherits="TyBus_Intranet_Test_V3.WorkReportOverTimes" %>

<asp:Content ID="WorkReportOverTimesForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">工作會報用加班時數統計</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalYM" runat="server" CssClass="text-Right-Blue" Text="計算年月：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCalYear" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:TextBox ID="eCalMonth" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbSplitMode" runat="server" CssClass="text-Right-Blue" Text="區分性別：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:RadioButtonList ID="rbSplitMode" runat="server" CssClass="text-Left-Black" RepeatColumns="2" OnSelectedIndexChanged="rbSplitMode_SelectedIndexChanged" Width="95%">
                        <asp:ListItem Text="不分性別" Value="0" Selected="True" />
                        <asp:ListItem Text="性別分別統計" Value="1" />
                    </asp:RadioButtonList>
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:TextBox ID="eSplitMode" runat="server" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="預覽資料" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="95%" />
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
    <asp:Panel ID="plShowData_Main0" runat="server" CssClass="ShowPanel">
        <asp:Panel ID="plShowData_Driver0" runat="server">
            <asp:GridView ID="gridShowData_Driver0" runat="server" BackColor="White" BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px" CellPadding="4" Width="30%" CssClass="SideBySide" AutoGenerateColumns="False" DataSourceID="sdsShowData_Driver0">
                <Columns>
                    <asp:BoundField DataField="DEPNO" HeaderText="單位" ReadOnly="True" SortExpression="DEPNO" />
                    <asp:BoundField DataField="DepName" HeaderText="列標籤" ReadOnly="True" SortExpression="DepName" />
                    <asp:BoundField DataField="TotalHR" HeaderText="時" ReadOnly="True" SortExpression="TotalHR" />
                    <asp:BoundField DataField="TotalMin" HeaderText="分" ReadOnly="True" SortExpression="TotalMin" />
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
        <asp:Panel ID="plShowData_Emp0" runat="server">
            <asp:GridView ID="gridShowData_Emp0" runat="server" BackColor="White" BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" CssClass="SideBySide" Width="65%">
                <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="#FFFFCC" />
                <PagerStyle BackColor="#FFFFCC" ForeColor="#330099" HorizontalAlign="Center" />
                <RowStyle BackColor="White" ForeColor="#330099" />
                <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="#663399" />
                <SortedAscendingCellStyle BackColor="#FEFCEB" />
                <SortedAscendingHeaderStyle BackColor="#AF0101" />
                <SortedDescendingCellStyle BackColor="#F6F0C0" />
                <SortedDescendingHeaderStyle BackColor="#7E0000" />
            </asp:GridView>
        </asp:Panel>
    </asp:Panel>

    <asp:Panel ID="plShowData_Main1" runat="server" CssClass="ShowPanel">
        <asp:Panel ID="plShowData_Driver1" runat="server">
            <asp:GridView ID="gridShowData_Driver1" runat="server" BackColor="White" BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px" CellPadding="4" Width="30%" CssClass="SideBySide" AutoGenerateColumns="False" DataSourceID="sdsShowData_Driver1">
                <Columns>
                    <asp:BoundField DataField="DepNo" HeaderText="單位" ReadOnly="True" SortExpression="DepNo" />
                    <asp:BoundField DataField="DepName" HeaderText="列標籤" ReadOnly="True" SortExpression="DepName" />
                    <asp:BoundField DataField="TotalHR_M" HeaderText="時 (男)" ReadOnly="True" SortExpression="TotalHR_M" />
                    <asp:BoundField DataField="TotalMin_M" HeaderText="分 (男)" ReadOnly="True" SortExpression="TotalMin_M" />
                    <asp:BoundField DataField="TotalHR_F" HeaderText="時 (女)" ReadOnly="True" SortExpression="TotalHR_F" />
                    <asp:BoundField DataField="TotalMin_F" HeaderText="分 (女)" ReadOnly="True" SortExpression="TotalMin_F" />
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
        <asp:Panel ID="plShowData_Emp1" runat="server">
            <asp:GridView ID="gridShowData_Emp1" runat="server" BackColor="White" BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" CssClass="SideBySide" Width="65%">
                <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="#FFFFCC" />
                <PagerStyle BackColor="#FFFFCC" ForeColor="#330099" HorizontalAlign="Center" />
                <RowStyle BackColor="White" ForeColor="#330099" />
                <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="#663399" />
                <SortedAscendingCellStyle BackColor="#FEFCEB" />
                <SortedAscendingHeaderStyle BackColor="#AF0101" />
                <SortedDescendingCellStyle BackColor="#F6F0C0" />
                <SortedDescendingHeaderStyle BackColor="#7E0000" />
            </asp:GridView>
        </asp:Panel>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsShowData_Driver0" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DEPNO)) AS DepName, CAST(SUM(TotalMin) AS int) / 60 AS TotalHR, CAST(SUM(TotalMin) AS int) % 60 AS TotalMin FROM (SELECT DEPNO, WORKHR * 60 + workmin - 480 AS TotalMin FROM RUNSHEETA WHERE (1 &lt;&gt; 1) UNION ALL SELECT DEPNO, WORKHR * 60 + workmin AS TotalMin FROM RUNSHEETA AS RUNSHEETA_1 WHERE (1 &lt;&gt; 1)) AS a GROUP BY DEPNO ORDER BY DEPNO"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsShowData_Driver1" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CAST('' AS varchar) AS DepNo, CAST('' AS varchar) AS DepName, CAST(0 AS int) AS TotalHR_M, CAST(0 AS int) AS TotalMin_M, CAST(0 AS int) AS TotalHR_F, CAST(0 AS int) AS TotalMin_F"></asp:SqlDataSource>
</asp:Content>
