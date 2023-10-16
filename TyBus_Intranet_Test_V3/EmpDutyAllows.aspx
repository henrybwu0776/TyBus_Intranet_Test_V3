<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmpDutyAllows.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmpDutyAllows" %>
<asp:Content ID="EmpDutyAllowsForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">員工補休申請審核</a>
    </div>
    <br />
    <div>
        <img alt="" src="img/sorry.gif" />
    </div>
    <asp:Panel ID="plSearch" runat="server" GroupingText="查詢條件" CssClass="SearchPanel" Visible="false">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="單位：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_S" runat="server" CssClass="text-Left-Black" AutoPostBack="true" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_E" runat="server" CssClass="text-Left-Black" AutoPostBack="true" Width="40%" />
                    <br />
                    <asp:Label ID="lbDepName_S" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="lbDepName_E" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbEmpNo_Search" runat="server" CssClass="text-Right-Blue" Text="申請人：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eEmpNo_S" runat="server" CssClass="text-Left-Black" AutoPostBack="true" Width="40%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eEmpNo_E" runat="server" CssClass="text-Left-Black" AutoPostBack="true" Width="40%" />
                    <br />
                    <asp:Label ID="lbEmpName_S" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="lbEmpName_E" runat="server" CssClass="text-Left-Black" Width="45%" />
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
