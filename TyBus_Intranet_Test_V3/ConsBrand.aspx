<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ConsBrand.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ConsBrand" %>

<asp:Content ID="ConsBrandForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">廠牌資料管理</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBrandCode_Search" runat="server" CssClass="text-Right-Blue" Text="廠牌代碼" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eBrandCode_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBrandName_Search" runat="server" CssClass="text-Right-Blue" Text="廠牌名稱" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eBrandName_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBelongGroup_Search" runat="server" CssClass="text-Right-Blue" Text="隸屬群組" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlBelongGroup_Search" runat="server" CssClass="text-Left-Black" Width="95%">
                        <asp:ListItem Text="" Value="" Selected="True" />
                        <asp:ListItem Text="通用" Value="00" />
                        <asp:ListItem Text="一般雜項類" Value="02" />
                        <asp:ListItem Text="電腦耗材類" Value="09" />
                    </asp:DropDownList>
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Blue" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" OnClick="bbExit_Click" Width="95%" />
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
    <asp:Panel ID="plShow" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridBrandList" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" AllowPaging="True" DataKeyNames="BrandCode" OnSelectedIndexChanged="gridBrandList_SelectedIndexChanged">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <Columns>
                <asp:ButtonField ButtonType="Button" Text="選擇" />
                <asp:BoundField DataField="BrandCode" HeaderText="廠牌代碼" />
                <asp:BoundField DataField="BrandName" HeaderText="廠牌名稱" />
                <asp:BoundField DataField="BelongGroup" HeaderText="隸屬群組" />
                <asp:BoundField DataField="Remark" HeaderText="備註" />
            </Columns>
            <EditRowStyle BackColor="#999999" />
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#E9E7E2" />
            <SortedAscendingHeaderStyle BackColor="#506C8C" />
            <SortedDescendingCellStyle BackColor="#FFFDF8" />
            <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
        </asp:GridView>
        <asp:Panel ID="plButton_List" runat="server" CssClass="ShowPanel-Detail_Float">
            <asp:Button ID="bbInsert" runat="server" CssClass="button-Black" Text="新增" Width="120px" OnClick="bbInsert_Click" />
            <asp:Button ID="bbEdit" runat="server" CssClass="button-Blue" Text="修改" Width="120px" OnClick="bbEdit_Click" />
            <asp:Button ID="bbDelete" runat="server" CssClass="button-Red" Text="刪除" Width="120px" OnClick="bbDelete_Click" />
        </asp:Panel>
        <asp:Panel ID="plButton_Action" runat="server" CssClass="ShowPanel-Detail_Float">
            <asp:Button ID="bbOK" runat="server" CssClass="button-Blue" Text="確定" Width="120px" OnClick="bbOK_Click" />
            <asp:Button ID="bbCancel" runat="server" CssClass="button-Red" Text="取消" Width="120px" OnClick="bbCancel_Click" />
        </asp:Panel>
        <asp:Label ID="lbErrorMessage" runat="server" CssClass="titleText-Red" />
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="lbBrandCode" runat="server" CssClass="text-Right-Blue" Text="廠牌名" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                    <asp:TextBox ID="eBrandCode" runat="server" CssClass="text-Left-Black" Text='<%# vdBrandCode %>' AutoPostBack="true" OnTextChanged="eBrandCode_TextChanged" Width="35%" />
                    <asp:TextBox ID="eBrandName" runat="server" CssClass="text-Left-Black" Text='<%# vdBrandName %>' Width="60%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="lbBelongGroup" runat="server" CssClass="text-Right-Blue" Text="隸屬群組" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:DropDownList ID="ddlBelongGroup" runat="server" CssClass="text-Left-Black" Width="95%">
                        <asp:ListItem Text="" Value="" Selected="True" />
                        <asp:ListItem Text="通用" Value="00" />
                        <asp:ListItem Text="一般雜項類" Value="02" />
                        <asp:ListItem Text="電腦耗材類" Value="09" />
                    </asp:DropDownList>
                    <asp:Label ID="eBelongGroup" runat="server" Text='<%# vdBelongGroup %>' Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="lbRemark" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                </td>
                <td class="ColBorder MultiLine_High ColWidth-5Col" colspan="4">
                    <asp:TextBox ID="eRemark" runat="server" TextMode="MultiLine" CssClass="text-Left-Black" Text='<%# vdRemark %>' Width="97%" Height="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="lbBuMan" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                    <asp:Label ID="eBuMan" runat="server" CssClass="text-Left-Black" Text='<%# vdBuMan %>' Width="35%" />
                    <asp:Label ID="eBuMan_C" runat="server" CssClass="text-Left-Black" Text='<%# vdBuMan_C %>' Width="60%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="lbBuDate" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="eBuDate" runat="server" CssClass="text-Left-Black" Text='<%# vdBuDate %>' Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="lbModifyMan" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                    <asp:Label ID="eModifyMan" runat="server" CssClass="text-Left-Black" Text='<%# vdModifyMan %>' Width="35%" />
                    <asp:Label ID="eModifyMan_C" runat="server" CssClass="text-Left-Black" Text='<%# vdModifyMan_C %>' Width="60%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="lbModifyDate" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-5Col">
                    <asp:Label ID="eModifyDate" runat="server" CssClass="text-Left-Black" Text='<%# vdModifyDate %>' Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-5Col" />
                <td class="ColWidth-5Col" />
                <td class="ColWidth-5Col" />
                <td class="ColWidth-5Col" />
                <td class="ColWidth-5Col" />
            </tr>
        </table>
    </asp:Panel>
</asp:Content>
