<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ExportRealCompanyPay.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ExportRealCompanyPay" %>
<asp:Content ID="ExportRealCompanyPayForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">匯出公司提撥金額</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <div>
            <asp:Label ID="lbTitle" runat="server" CssClass="text-Left-Black" Text="匯出公司提撥金額" />
        </div>
        <br />
        <div>
            <asp:Label ID="lbCalYM" runat="server" CssClass="text-Left-Black" Text="查核年月 (西元年 YYYY/MM) " />
            <asp:TextBox runat="server" ID="eCalYM" CssClass="text-Left-Black" />
        </div>
        <br />
        <div>
            <asp:Button ID="bbOK" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbOK_Click" Width="100px" />
            <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="100px" />
        </div>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:Button ID="bbExcel" runat="server" CssClass="button-Red" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="120px" />
        <asp:GridView ID="gridShowData" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px" CellPadding="4" DataSourceID="sdsShowData" Width="100%">
            <Columns>
                <asp:BoundField DataField="depno" HeaderText="部門編號" SortExpression="depno" />
                <asp:BoundField DataField="DepName" HeaderText="部門名稱" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="empno" HeaderText="員工編號" SortExpression="empno" />
                <asp:BoundField DataField="name" HeaderText="姓名" SortExpression="name" />
                <asp:BoundField DataField="paydate" DataFormatString="{0:D}" HeaderText="計薪年月" SortExpression="paydate" />
                <asp:BoundField DataField="cashnum71" HeaderText="勞保公司負擔" SortExpression="cashnum71" />
                <asp:BoundField DataField="cashnum72" HeaderText="健保公司負擔" SortExpression="cashnum72" />
                <asp:BoundField DataField="cashnum74" HeaderText="勞退公司提撥" SortExpression="cashnum74" />
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
        <asp:SqlDataSource ID="sdsShowData" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT depno, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.depno)) AS DepName, empno, name, paydate, cashnum71, cashnum72, cashnum74 FROM PAYREC AS e WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    </asp:Panel>
</asp:Content>
