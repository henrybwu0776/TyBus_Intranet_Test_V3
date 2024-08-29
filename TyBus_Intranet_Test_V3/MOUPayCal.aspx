<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="MOUPayCal.aspx.cs" Inherits="TyBus_Intranet_Test_V3.MOUPayCal" %>

<asp:Content ID="MOUPayCalForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">MOD專案駕駛薪資計算</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColWidth-7Col" rowspan="7"/>
                <td class="ColHeight ColBorder ColWidth-7Col" rowspan="7">
                    <asp:ListBox ID="eDriverList" runat="server" CssClass="text-Left-Black" Width="95%" SelectionMode="Multiple" Height="95%" />
                </td>
                <td class="ColHeight ColWidth-7Col" rowspan="7">
                    <asp:Button ID="bbSelectTo" runat="server" CssClass="button-Black" Text="-->" OnClick="bbSelectTo_Click" Width="120px" />
                    <br />
                    <asp:Button ID="bbSelectAll" runat="server" CssClass="button-Blue" Text="全選" OnClick="bbSelectAll_Click" Width="120px" />
                    <br />
                    <asp:Button ID="bbUnselectTo" runat="server" CssClass="button-Red" Text="<--" OnClick="bbUnselectTo_Click" Width="120px" />
                    <br />
                    <asp:Button ID="bbUnselectAll" runat="server" CssClass="button-Red" Text="全部不選" OnClick="bbUnselectAll_Click" Width="120px" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" rowspan="7">
                    <asp:ListBox ID="eSelectedList" runat="server" CssClass="text-Left-Black" SelectionMode="Multiple" Width="95%" Height="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbCalYM" runat="server" CssClass="text-Right-Blue" Text="計算年月(民國年)" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eCalYear" runat="server" CssClass="text-Left-Black" Width="25%" />
                    <asp:Label ID="lbCalYear" runat="server" CssClass="text-Left-Black" Text="年" />
                    <asp:TextBox ID="eCalMonth" runat="server" CssClass="text-Left-Black" Width="25%" />
                    <asp:Label ID="lbCalMonth" runat="server" CssClass="text-Left-Black" Text="月" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbPayAmount" runat="server" CssClass="text-Right-Blue" Text="薪資總額" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="ePayAmount" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="ePayAmount_TextChanged" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbGovPayRatio" runat="server" CssClass="text-Right-Blue" Text="政府補助比例" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eGovPayRatio" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="ePayAmount_TextChanged" Width="80%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="％" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbCompanyPay" runat="server" CssClass="text-Right-Blue" Text="公司負擔金額" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eCompanyPay" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eCompanyPay_TextChanged" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbGovPayAMT" runat="server" CssClass="text-Right-Blue" Text="政府補助金額" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="eGovPayAMT" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbBonusCode" runat="server" CssClass="text-Right-Blue" Text="津貼項目代碼" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eBonusCode" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-7Col" colspan="3">
                    <asp:Button ID="bbCalGreatBounds" runat="server" CssClass="button-Blue" Text="計算里程碑獎金" OnClick="bbCalGreatBounds_Click" Width="120px" />
                    <asp:Button ID="bbCal" runat="server" CssClass="button-Black" Text="計算" OnClick="bbCal_Click" Width="120px" />
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Blue" Text="查詢" OnClick="bbSearch_Click" Width="120px" />
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" OnClick="bbExit_Click" Width="120px" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-7Col" />
                <td class="ColHeight ColWidth-7Col" />
                <td class="ColHeight ColWidth-7Col" />
                <td class="ColHeight ColWidth-7Col" />
                <td class="ColHeight ColWidth-7Col" />
                <td class="ColHeight ColWidth-7Col" />
                <td class="ColHeight ColWidth-7Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gvDataList" runat="server" Width="100%" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" GridLines="None">
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
