<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ConsumablesInstore.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ConsumablesInstore" %>

<asp:Content ID="ConsumablesInstoreForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">總務耗材進貨</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbSheetNo_Search" runat="server" CssClass="text-Right-Blue" Text="進貨單號" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eSheetNo_S_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="titleText-Black" Text="~" />
                    <asp:TextBox ID="eSheetNo_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbSupNo_Search" runat="server" CssClass="text-Right-Blue" Text="出貨廠商" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eSupNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eSupNo_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eSupName_Search" runat="server" CssClass="text-Left-Black" Width="60%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbSheetStatus_Search" runat="server" CssClass="text-Right-Blue" Text="進貨狀況" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlSheetStatus_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="True" OnSelectedIndexChanged="ddlSheetStatus_Search_SelectedIndexChanged" Width="95%" />
                    <asp:Label ID="eSheetStatus_Search" runat="server" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuMan_Search" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eBuMan_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eBuMan_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eBuManName_Search" runat="server" CssClass="text-Left-Black" Width="60%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuDate_Search" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eBuDate_S_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="titleText-Black" Text="~" />
                    <asp:TextBox ID="eBuDate_E_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbSheetNote_Search" runat="server" CssClass="text-Right-Blue" Text="單據摘要" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eSheetNote_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbConsNo_Search" runat="server" CssClass="text-Right-Blue" Text="料號" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eConsNo_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbConsName_Search" runat="server" CssClass="text-Right-Blue" Text="品名" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eConsName_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBrand_Search" runat="server" CssClass="text-Right-Blue" Text="廠牌" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlBrand_Search" runat="server" CssClass="text-Left-Black" Width="95%" AutoPostBack="True" OnSelectedIndexChanged="ddlBrand_Search_SelectedIndexChanged" />
                    <asp:Label ID="eBrand_Search" runat="server" Visible="false" />
                </td>
                <td class="ColWidth-8Col" colspan="2" />
            </tr>
            <tr>
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" colspan="2">
                    <asp:Button ID="bbOK_Search" runat="server" CssClass="button-Blue" OnClick="bbOK_Search_Click" Text="查詢" Width="120px" />
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" OnClick="bbClose_Click" Text="離開" Width="120px" />
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
    <asp:Panel ID="plShowA" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridSheetA" runat="server" Width="100%" BackColor="White" BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" CellPadding="3" GridLines="Horizontal" AutoGenerateColumns="False">
            <AlternatingRowStyle BackColor="#F7F7F7" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="SheetNo" HeaderText="單號" ReadOnly="True" SortExpression="SheetNo" />
                <asp:BoundField DataField="SupNo" HeaderText="SupNo" SortExpression="SupNo" Visible="False" />
                <asp:BoundField DataField="SupName" HeaderText="供應商" ReadOnly="True" SortExpression="SupName" />
                <asp:BoundField DataField="BuDate" DataFormatString="{0:D}" HeaderText="開單日期" SortExpression="BuDate" />
                <asp:BoundField DataField="TotalAmount" HeaderText="單據總金額" SortExpression="TotalAmount" />
                <asp:BoundField DataField="SheetNote" HeaderText="摘要" SortExpression="SheetNote" />
                <asp:BoundField DataField="SheetStatus_C" HeaderText="執行狀態" ReadOnly="True" SortExpression="SheetStatus_C" />
                <asp:BoundField DataField="RemarkA" HeaderText="備註" SortExpression="RemarkA" />
            </Columns>
            <FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
            <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#F7F7F7" />
            <PagerStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" HorizontalAlign="Right" />
            <RowStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="#F7F7F7" />
            <SortedAscendingCellStyle BackColor="#F4F4FD" />
            <SortedAscendingHeaderStyle BackColor="#5A4C9D" />
            <SortedDescendingCellStyle BackColor="#D8D8F0" />
            <SortedDescendingHeaderStyle BackColor="#3E3277" />
        </asp:GridView>
        <asp:Panel ID="plShowDetailA" runat="server" CssClass="ShowPanel">
            <asp:Panel ID="plButton_List" runat="server" CssClass="ShowPanel-Detail_Float">
                <asp:Button ID="bbInsert" runat="server" CssClass="button-Black" Text="新增" OnClick="bbInsert_Click" Width="120px" />
                <asp:Button ID="bbEdit" runat="server" CssClass="button-Blue" Text="修改" OnClick="bbEdit_Click" Width="120px" />
                <asp:Button ID="bbDelete" runat="server" CssClass="button-Red" Text="刪除" OnClick="bbDelete_Click" Width="120px" />
            </asp:Panel>
            <asp:Panel ID="plButton_Action" runat="server" CssClass="ShowPanel-Detail_Float">
                <asp:Button ID="bbOK" runat="server" CssClass="button-Blue" Text="確定" OnClick="bbOK_Click" Width="120px" />
                <asp:Button ID="bbCancel" runat="server" CssClass="button-Red" Text="取消" OnClick="bbCancel_Click" Width="120px" />
            </asp:Panel>
            <asp:Panel ID="plButton_ChangeMode" runat="server" CssClass="ShowPanel-Detail_Float">
                <asp:Button ID="bbIsArrived" runat="server" CssClass="button-Black" Text="收貨" OnClick="bbGetGoods_Click" Width="120px" />
                <asp:Button ID="bbIsPaid" runat="server" CssClass="button-Blue" Text="付款" OnClick="bbIsPaid_Click" Width="120px" />
                <asp:Button ID="bbIsAbort" runat="server" CssClass="button-Red" Text="作廢" OnClick="bbIsAbort_Click" Width="120px" />
                <asp:Button ID="bbCaseClose" runat="server" CssClass="button-Blue" Text="結案" OnClick="bbCaseClose_Click" Width="120px" />
            </asp:Panel>
            <table class="TableSetting">
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbSheetNo" runat="server" CssClass="text-Right-Blue" Text="進貨單號" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eSheetNo" runat="server" CssClass="text-Left-Black" Text='<%# vdSheetNo %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbSubName" runat="server" CssClass="text-Right-Blue" Text="供應商" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:TextBox ID="eSupNo" runat="server" CssClass="text-Left-Black" Text='<%# vdSupNo %>' OnTextChanged="eSupNo_TextChanged" Width="35%" />
                        <asp:Label ID="eSupName" runat="server" CssClass="text-Left-Black" Text='<%# vdSupName %>' Width="55%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbSheetStatus" runat="server" CssClass="text-Right-Blue" Text="單據狀態" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eSheetStatus" runat="server" Text='<%#vdSheetStatus %>' Visible="false" />
                        <asp:Label ID="eSheetStatus_C" runat="server" CssClass="text-Left-Black" Text='<%# vdSheetStatus_C %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbStatusDate" runat="server" CssClass="text-Right-Blue" Text="狀態日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eStatusDate" runat="server" CssClass="text-Left-Black" Text='<%# vdStatusDate %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbBuDate" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eBuDate" runat="server" CssClass="text-Left-Black" Text='<%# vdBuDate %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbBuMan" runat="server" CssClass="text-Right-Blue" Text="開單人員" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:TextBox ID="eBuMan" runat="server" CssClass="text-Left-Black" Text='<%# vdBuMan %>' OnTextChanged="eBuMan_TextChanged" Width="35%" />
                        <asp:Label ID="eBuManName" runat="server" CssClass="text-Left-Black" Text='<%# vdBuManName %>' Width="55%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbPayMode" runat="server" CssClass="text-Right-Blue" Text="付款方式" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:DropDownList ID="ddlPayMode" runat="server" CssClass="text-Left-Black" Width="95%" OnSelectedIndexChanged="ddlPayMode_SelectedIndexChanged" />
                        <asp:Label ID="ePayMode" runat="server" Text='<%# vdPayMode %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbPayDate" runat="server" CssClass="text-Right-Blue" Text="付款日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="ePayDate" runat="server" CssClass="text-Left-Black" Text='<%# vdPayDate %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbSheetNote" runat="server" CssClass="text-Right-Blue" Text="摘要" Width="95%" />
                    </td>
                    <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="9">
                        <asp:TextBox ID="eSheetNote" runat="server" CssClass="MultiLine_High" Text='<%# vdSheetNote %>' TextMode="MultiLine" Width="97%" Height="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbAmount" runat="server" CssClass="text-Right-Blue" Text="小計金額" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eAmount" runat="server" CssClass="text-Left-Black" Text='<%# vdAmount %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbTaxType" runat="server" CssClass="text-Right-Blue" Text="稅別" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eTaxType" runat="server" CssClass="text-Left-Black" Text='<%# vdTaxType_C %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbTaxRate" runat="server" CssClass="text-Right-Blue" Text="稅率" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eTaxRate" runat="server" CssClass="text-Left-Black" Text='<%# vdTaxRate %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbTaxAMT" runat="server" CssClass="text-Right-Blue" Text="稅額" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eTaxAMT" runat="server" CssClass="text-Left-Black" Text='<%# vdTaxAMT %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbTotalAmount" runat="server" CssClass="text-Right-Blue" Text="總金額" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eTotalAmount" runat="server" CssClass="text-Left-Black" Text='<%# vdTotalAmount %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbRemarkA" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                    </td>
                    <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="9">
                        <asp:TextBox ID="eRemarkA" runat="server" CssClass="text-Left-Black" Text='<%# vdRemarkA %>' TextMode="MultiLine" Width="97%" Height="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColWidth-10Col" />
                    <td class="ColWidth-10Col" />
                    <td class="ColWidth-10Col" />
                    <td class="ColWidth-10Col" />
                    <td class="ColWidth-10Col" />
                    <td class="ColWidth-10Col" />
                    <td class="ColWidth-10Col" />
                    <td class="ColWidth-10Col" />
                    <td class="ColWidth-10Col" />
                    <td class="ColWidth-10Col" />
                </tr>
            </table>
        </asp:Panel>

    </asp:Panel>
</asp:Content>
