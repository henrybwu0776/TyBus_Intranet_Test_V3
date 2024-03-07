<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmpHealthCheckList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmpHealthCheckList" %>
<asp:Content ID="EmpHealthCheckListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">年度健康檢查名冊</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbCalBaseDate" runat="server" CssClass="text-Right-Blue" Text="結算日期" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eCalBaseDate" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出EXCEL" Width="95%" OnClick="bbExcel_Click" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" Width="95%" OnClick="bbExit_Click" />
                </td>
                <td class="ColHeight ColWidth-7Col" colspan="3">
                    <asp:Label ID="lbErrorMSG" runat="server" CssClass="text-Left-Red" Text="" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
            </tr>
        </table>
    </asp:Panel>
</asp:Content>
