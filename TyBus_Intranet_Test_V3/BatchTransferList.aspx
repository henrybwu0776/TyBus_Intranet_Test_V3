<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="BatchTransferList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.BatchTransferList" %>

<asp:Content ID="BatchTransferListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">整批匯款名冊轉換</a>
    </div>
    <br />
    <asp:Panel ID="plMain" runat="server" CssClass="ShowPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbPayAccount" runat="server" CssClass="text-Right-Blue" Text="付款帳號：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="ePayAccount" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbSourceTXT" runat="server" CssClass="text-Right-Blue" Text="來源TXT檔：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="4">
                    <asp:FileUpload ID="fuSourceTXT" runat="server" Width="90%" CssClass="text-Left-Black" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbTrans" runat="server" CssClass="button-Black" Text="開始轉檔" OnClick="bbTrans_Click" Width="90%" UseSubmitBehavior="False" OnClientClick="this.disabled=true;" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
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
</asp:Content>
