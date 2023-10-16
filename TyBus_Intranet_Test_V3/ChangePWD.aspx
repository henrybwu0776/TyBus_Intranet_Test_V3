<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ChangePWD.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ChangePWD" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">變更密碼</a>
    </div>
    <br />
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShow" runat="server" CssClass="PWDPanel">
        <br />
        <br />
        <asp:Label ID="lbLoginID" runat="server" CssClass="text-Left-Black" Text="使用者帳號：" width="120px"/>
        <asp:Label ID="eLoginID" runat="server" CssClass="text-Left-Blue" />
        <br />
        <br />
        <asp:Label ID="lbChangePWD" runat="server" CssClass="text-Left-Black" Text="新密碼：" Width="120px" />
        <asp:TextBox ID="eChangePWD" runat="server" CssClass="text-Left-Blue" Width="200px" TextMode="Password" />
        <br />
        <br />
        <asp:Label ID="lbConfirmPWD" runat="server" CssClass="text-Left-Black" Text="確認新密碼：" Width="120px" />
        <asp:TextBox ID="eConfirmPWD" runat="server" CssClass="text-Left-Blue" Width="200px" TextMode="Password" />
        <br />
        <br />
        <asp:Button ID="bbOK" runat="server" Text="確定" CssClass="button-Blue" OnClick="bbOK_Click" Width="120px" UseSubmitBehavior="False" OnClientClick="this.disabled=true;" />
        <asp:Button ID="bbExit" runat="server" Text="放棄" CssClass="button-Red" OnClick="bbExit_Click" Width="120px" />
        </asp:Panel>
    <asp:CompareValidator ID="CompareValidator1" runat="server" ControlToCompare="eChangePWD" ControlToValidate="eConfirmPWD" CssClass="text-Left-Red"></asp:CompareValidator>
</asp:Content>
