<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CarInspectionDataCheck.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CarInspectionDataCheck" %>

<asp:Content ID="CarInspectionDataCheckForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">車輛檢驗資料查詢</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCarID_Search" runat="server" CssClass="text-Right-Blue" Text="牌照號碼：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCarID_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNoS_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNoS_Search_TextChanged" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eDepNoE_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNoE_Search_TextChanged" Width="40%" />
                    <br>
                    <asp:Label ID="eDepNameS_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="eDepNameE_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCheckDate_Search" runat="server" CssClass="text-Right-Blue" Text="完成檢驗日：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eCheckDateS_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eCheckDateE_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbShowData" runat="server" CssClass="button-Black" OnClick="bbShowData_Click" Text="查詢" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" OnClick="bbExcel_Click" Text="匯出 EXCEL" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" OnClick="bbClose_Click" Text="結束" Width="90%" />
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
        <asp:GridView ID="gridShowData" runat="server" AutoGenerateColumns="False" CellPadding="4" DataSourceID="sdsShowData" ForeColor="#333333" GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <Columns>
                <asp:BoundField DataField="Car_ID" HeaderText="牌照號碼" SortExpression="Car_ID" />
                <asp:BoundField DataField="GDate" DataFormatString="{0:d}" HeaderText="領照日期" SortExpression="GDate" />
                <asp:BoundField DataField="bespeakdate" DataFormatString="{0:d}" HeaderText="預定檢驗日" SortExpression="bespeakdate" />
                <asp:BoundField DataField="CheckDate" DataFormatString="{0:d}" HeaderText="完成檢驗日" SortExpression="CheckDate" />
                <asp:BoundField DataField="nextedate" DataFormatString="{0:d}" HeaderText="下次檢驗日" SortExpression="nextedate" />
                <asp:BoundField DataField="NCheckTerm" DataFormatString="{0:d}" HeaderText="下次檢驗期限" SortExpression="NCheckTerm" />
                <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="StoreName" HeaderText="檢驗廠商" ReadOnly="True" SortExpression="StoreName" />
            </Columns>
            <EditRowStyle BackColor="#999999" />
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#E9E7E2" />
            <SortedAscendingHeaderStyle BackColor="#506C8C" />
            <SortedDescendingCellStyle BackColor="#FFFDF8" />
            <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
        </asp:GridView>
        <asp:SqlDataSource ID="sdsShowData" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT c.Car_ID, c.GDate, c.bespeakdate, c.CheckDate, a.nextedate, a.NCheckTerm, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.CompanyNo)) AS DepName, (SELECT name FROM CUSTOM WHERE (types = 'S') AND (code = c.storeno)) AS StoreName FROM Car_infoC AS c LEFT OUTER JOIN Car_infoA AS a ON a.Car_ID = c.Car_ID WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    </asp:Panel>
</asp:Content>
