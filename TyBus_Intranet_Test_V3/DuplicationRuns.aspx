<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DuplicationRuns.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DuplicationRuns" %>

<asp:Content ID="DuplicationRunsForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">行車班次表複製</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eDepName_Search" runat="server" CssClass="text-Left-Black" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbRunturn_Search" runat="server" CssClass="text-Right-Blue" Text="班次" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlRunturn_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col"></td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Blue" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" OnClick="bbExit_Click" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShow" runat="server" CssClass="ShowPanel">
        <div class="ShowPanel-Detail-TwoColumns ColBorder">
            <table class="TableSetting">
                <tr>
                    <td class="ColHeight ColWidth-4Col" />
                    <td class="ColHeight ColWidth-4Col" />
                    <td class="ColHeight ColWidth-4Col" />
                    <td class="ColHeight ColWidth-4Col" />
                </tr>
            </table>
        </div>
        <div class="ShowPanel-Detail-TwoColumns ColBorder">
            <table class="TableSetting">
                <tr>
                    <td class="ColHeight ColWidth-4Col" />
                    <td class="ColHeight ColWidth-4Col" />
                    <td class="ColHeight ColWidth-4Col" />
                    <td class="ColHeight ColWidth-4Col" />
                </tr>
            </table>
        </div>
    </asp:Panel>
</asp:Content>
