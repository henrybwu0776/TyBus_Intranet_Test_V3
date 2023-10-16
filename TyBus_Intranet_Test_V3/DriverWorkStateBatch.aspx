<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DriverWorkStateBatch.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DriverWorkStateBatch" %>
<asp:Content ID="DriverWorkStateBatchForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">駕駛員例休批次修改</a>
    </div>
    <br />
    <asp:Panel ID="plShow" runat="server" CssClass="ShowPanel" width="100%">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Right-Black" OnTextChanged="eDepNo_Edit_TextChanged" Width="30%" AutoPostBack="True" />
                    <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Width="60%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbSheetDate_Edit" runat="server" CssClass="text-Right-Blue" Text="憑單日期：" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eSheetDate_Edit" runat="server" CssClass="text-Left-Black" Width="50%" />
                    <asp:Button ID="bbGetDriverList" runat="server" CssClass="button-Red" Text="重整" OnClick="bbGetDriverList_Click" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Label ID="lbDriver_Edit" runat="server" CssClass="titleText-S-Blue" Text="可選駕駛員" Width="100%" />
                </td>
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col">
                    <asp:Label ID="lbCheckedDirver_Edit" runat="server" CssClass="titleText-S-Blue" Text="待調整駕駛員" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:Label ID="lbWorkState_Edit" runat="server" CssClass="titleText-S-Blue" Text="出勤狀況" Width="100%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col" rowspan="10">
                    <asp:ListBox ID="eDriverList_Edit" runat="server" CssClass="text-Left-Blue" Width="100%" Height="100%" SelectionMode="Multiple" />
                </td>
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColBorder ColWidth-6Col" rowspan="10">
                    <asp:ListBox ID="eFixedDriver_Edit" runat="server" CssClass="text-Left-Black" Width="100%" Height="100%" SelectionMode="Multiple" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" rowspan="2" colspan="2">
                    <asp:RadioButtonList ID="eWorkState_Edit" runat="server" CssClass="text-Left-Black" Width="100%" Height="100%" RepeatColumns="4">
                        <asp:ListItem Text="平日" Value="0" />
                        <asp:ListItem Text="休假" Value="1" />
                        <asp:ListItem Text="例假日" Value="2" />
                        <asp:ListItem Text="國定例假" Value="3" />
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" />
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbOK_Edit" runat="server" CssClass="button-Blue" Text="確定" OnClick="bbOK_Edit_Click" Width="40%" />
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="離開" OnClick="bbClose_Click" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" />
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbGetDriver_Edit" runat="server" CssClass="button-Blue" Width="100%" OnClick="bbGetDriver_Edit_Click" Text="選取-->" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbCancelDriver_Edit" runat="server" CssClass="button-Red" Width="100%" OnClick="bbCancelDriver_Edit_Click" Text="<--取消" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" />
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" />
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" />
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" />
            </tr>
            <tr>
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
            </tr>
        </table>
    </asp:Panel>
</asp:Content>
