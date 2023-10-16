<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CalCompanyPay.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CalCompanyPay" %>

<asp:Content ID="CalCompanyPayForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">重新計算公司提撥</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <div>
            <asp:Label ID="lbTitle" runat="server" CssClass="text-Left-Black" Text="重新計算公司負擔金額" />
        </div>
        <br />
        <div>
            <asp:Label ID="lbCalYM" runat="server" CssClass="text-Left-Black" Text="計算年月 (西元年 YYYY/MM) " />
            <asp:TextBox runat="server" ID="eCalYM" CssClass="text-Left-Black" />
        </div>
        <br />
        <div>
            <asp:Button ID="bbOK" runat="server" CssClass="button-Black" Text="計算" OnClick="bbOK_Click" Width="100px" OnClientClick="this.disabled=true;" UseSubmitBehavior="False" />
            <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="100px" />
        </div>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel" Visible="false">
        <asp:Button runat="server" ID="bbExportExcel" CssClass="button-Blue" Text="匯出EXCEL" OnClick="bbExportExcel_Click" />
        <asp:GridView ID="gridData" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" GridLines="None">
            <Columns>
                <asp:BoundField DataField="DepNo" HeaderText="部門編號" />
                <asp:BoundField DataField="DepName" HeaderText="部門名稱" />
                <asp:BoundField DataField="EmpNo" HeaderText="員工編號" />
                <asp:BoundField DataField="EmpName" HeaderText="姓名" />
                <asp:BoundField DataField="CalDate" HeaderText="計薪年月" DataFormatString="{0:d}" />
                <asp:BoundField DataField="LaiCompanyAMT" HeaderText="勞保公司負擔" />
                <asp:BoundField DataField="HiCompanyAMT" HeaderText="健保公司負擔" />
                <asp:BoundField DataField="LaiRetireAMT" HeaderText="勞退公司提撥" />
                <asp:BoundField DataField="LaiSafeAMT" HeaderText="勞保墊償基金提撥" />
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
</asp:Content>
