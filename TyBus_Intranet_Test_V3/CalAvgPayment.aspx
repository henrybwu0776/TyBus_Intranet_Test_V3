<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CalAvgPayment.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CalAvgPayment" %>

<asp:Content ID="CalAvgPaymentForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">計算平均薪資</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <div>
            <asp:Label ID="lbTitle" runat="server" CssClass="text-Left-Black" Text="計算平均薪資" />
            <br />
            <asp:Label ID="lbAlert" runat="server" CssClass="text-Left-Red" Text="***每年更新勞健保級距時請通知電腦課維護 {BLOCK} 欄位的值***" />
            <br />
            <asp:Button ID="bbOK" runat="server" CssClass="button-Black" Text="資料匯出" OnClick="bbOK_Click" Width="120px" />
            <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="120px" />
        </div>
    </asp:Panel>
</asp:Content>
