<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="RunSheetCheck.aspx.cs" Inherits="TyBus_Intranet_Test_V3.RunSheetCheck" %>

<asp:Content ID="RunSheetCheckForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">行車憑單錯誤檢查</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_S" runat="server" CssClass="text-Left-Black" OnTextChanged="eDepNo_S_TextChanged" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_E" runat="server" CssClass="text-Left-Black" OnTextChanged="eDepNo_E_TextChanged" Width="40%" />
                    <br />
                    <asp:Label ID="eDepName_S" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="eDepName_E" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCheckYM" runat="server" CssClass="text-Right-Blue" Text="檢查年月：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="4">
                    <asp:TextBox ID="eCheckYear" runat="server" CssClass="text-Left-Black" Width="120px" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text=" 年 " Width="20px" />
                    <asp:TextBox ID="eCheckMonth" runat="server" CssClass="text-Left-Black" Width="100px" />
                    <asp:Label ID="lbSplit_3" runat="server" CssClass="text-Left-Black" Text=" 月" Width="20px" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" OnClick="bbSearch_Click" Text="查詢" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Red" OnClick="bbExcel_Click" Text="匯出 EXCEL" Width="95%" />
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
        <asp:GridView ID="gridRunSheetCheckList" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="ASSIGNNO" DataSourceID="sdsRunSheetCheckList" GridLines="None" Width="100%">
            <Columns>
                <asp:BoundField DataField="ASSIGNNO" HeaderText="憑單單號" ReadOnly="True" SortExpression="ASSIGNNO" />
                <asp:BoundField DataField="BUDATE" HeaderText="行車日期" SortExpression="BUDATE" DataFormatString="{0:d}" />
                <asp:BoundField DataField="WorkState" HeaderText="出勤狀況" SortExpression="WorkState" />
                <asp:BoundField DataField="DEPNO" HeaderText="站別編號" SortExpression="DEPNO" />
                <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="DRIVER" HeaderText="駕駛員工號" SortExpression="DRIVER" />
                <asp:BoundField DataField="DriverName" HeaderText="姓名" ReadOnly="True" SortExpression="DriverName" />
                <asp:BoundField DataField="WORKHR" HeaderText="上班小時" SortExpression="WORKHR" />
                <asp:BoundField DataField="workmin" HeaderText="上班分鐘" SortExpression="workmin" />
                <asp:BoundField DataField="ActualKm" HeaderText="總公里數" SortExpression="ActualKm" />
                <asp:BoundField DataField="HasESCDuty" HeaderText="當日有請假單" SortExpression="HasESCDuty" />
                <asp:BoundField DataField="Remark" HeaderText="其他異常情況" SortExpression="Remark" />
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
    <asp:SqlDataSource ID="sdsRunSheetCheckList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT ASSIGNNO, BUDATE, DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = RUNSHEETA.DEPNO)) AS DepName, DRIVER, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = RUNSHEETA.DRIVER)) AS DriverName, WORKHR, workmin, ActualKm, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = RUNSHEETA.WORKSTATE) AND (FKEY = '行車記錄單A     runsheeta       WORKSTATE')) AS WorkState, CASE WHEN (SELECT COUNT(SerialNo) FROM ESCDuty WHERE ESCDUty.RealDay = RunSheetA.BuDate AND ESCDuty.ApplyMan = RunSheetA.Driver) = 0 THEN '' ELSE 'V' END AS HasESCDuty, CAST('' AS varchar) AS Remark FROM RUNSHEETA WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
