<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CreateIAAsset.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CreateIAAsset" %>

<asp:Content ID="IAAsetForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:Panel runat="server" ID="plLoginData" CssClass="PanelMargin">
        <a class="a1">登入人員：</a>
        <asp:Label CssClass="a1" ID="lbLoginID" runat="server" Text="" />
        <asp:Label CssClass="a1" ID="lbLoginName" runat="server" Text="" />
        <asp:Label CssClass="a1" ID="lbSQLStr" runat="server" Text="" Visible="false" />
    </asp:Panel>
    <asp:Panel runat="server" ID="plSearch" GroupingText="查詢條件" CssClass="PanelMargin">
        <table class="TableSetting">
            <tr>
                <td class="ColWidth2 ColHeight ColBorder a3">
                    <asp:CheckBox ID="cbAssetType" runat="server" Text="類別：" />
                    <asp:DropDownList ID="ddlAssetTypeSearch" runat="server" OnSelectedIndexChanged="ddlAssetTypeSearch_SelectedIndexChanged" AutoPostBack="True">
                        <asp:ListItem Value="S" Text="軟體" Selected="True" />
                        <asp:ListItem Value="H" Text="硬體" />
                        <asp:ListItem Value="C" Text="耗材" />
                    </asp:DropDownList>
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">
                    <asp:CheckBox ID="cbSystematics" runat="server" Text="品項：" />
                    <asp:DropDownList ID="ddlSystematicsSearch" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">
                    <asp:CheckBox ID="cbBrand" runat="server" Text="廠牌：" />
                    <asp:DropDownList ID="ddlBrandSearch" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">品名：
                    <asp:TextBox ID="eProductName_Search" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3" rowspan="3">
                    <asp:Button ID="bbSearch" runat="server" Text="搜尋資產" CssClass="b3" OnClick="bbSearch_Click" />
                    <asp:Button ID="bbClearSearch" runat="server" Text="清除重查" CssClass="b4" OnClick="bbClearSearch_Click" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth2 ColHeight ColBorder a3" colspan="2">資產編號：
                    <asp:TextBox ID="eAssetNo_Search1" runat="server" />
                    <asp:Label ID="lbSplit1" runat="server" Text=" ～ " />
                    <asp:TextBox ID="eAssetNo_Search2" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3" colspan="2">購置日期：
                    <asp:TextBox ID="eBuyDate_Search1" runat="server" />
                    <asp:Label ID="lbSplit2" runat="server" Text=" ～ " />
                    <asp:TextBox ID="eBuyDate_Search2" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth2 ColHeight ColBorder a3" colspan="2">供應商：
                    <asp:TextBox ID="eSupplier_Search1" runat="server" />
                    <asp:Label ID="lbSplit3" runat="server" Text=" ～ " />
                    <asp:TextBox ID="eSupplier_Search2" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3" colspan="2">舊編號：
                    <asp:TextBox ID="eOldAssetNo_Search1" runat="server" />
                    <asp:Label ID="lbSplit4" runat="server" Text=" ～ " />
                    <asp:TextBox ID="eOldAssetNo_Search2" runat="server" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel runat="server" ID="plDataGrid">
        <asp:GridView runat="server" ID="gvAssetList" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="AssetNo" DataSourceID="dsAssetList" GridLines="None" Width="100%" OnSelectedIndexChanged="gvAssetData_SelectedIndexChanged">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="AssetNo" HeaderText="資產編號" ReadOnly="True" SortExpression="AssetNo" />
                <asp:BoundField DataField="AssetType_C" HeaderText="類別" ReadOnly="True" SortExpression="AssetType_C" />
                <asp:BoundField DataField="AssetType" HeaderText="AssetType" SortExpression="AssetType" Visible="False" />
                <asp:BoundField DataField="Systematics_C" HeaderText="品項" ReadOnly="True" SortExpression="Systematics_C" />
                <asp:BoundField DataField="Systematics" HeaderText="Systematics" SortExpression="Systematics" Visible="False" />
                <asp:BoundField DataField="Brand_C" HeaderText="廠牌" ReadOnly="True" SortExpression="Brand_C" />
                <asp:BoundField DataField="Brand" HeaderText="Brand" SortExpression="Brand" Visible="False" />
                <asp:BoundField DataField="OriModeNumber" HeaderText="原廠型號" SortExpression="OriModeNumber" />
                <asp:BoundField DataField="ProductName" HeaderText="品名" SortExpression="ProductName" />
                <asp:BoundField DataField="OriSerialNumber" HeaderText="原廠序號" SortExpression="OriSerialNumber" />
                <asp:BoundField DataField="BuyDate" HeaderText="購置日期" ReadOnly="True" SortExpression="BuyDate" />
                <asp:BoundField DataField="Warranty" HeaderText="保固年限" SortExpression="Warranty" />
                <asp:BoundField DataField="Supplier_C" HeaderText="供應商" ReadOnly="True" SortExpression="Supplier_C" />
                <asp:BoundField DataField="Supplier" HeaderText="Supplier" SortExpression="Supplier" Visible="False" />
                <asp:BoundField DataField="WarrantyDate" HeaderText="保固到期日" ReadOnly="True" SortExpression="WarrantyDate" />
                <asp:BoundField DataField="OldAssetNo" HeaderText="舊資產編號" SortExpression="OldAssetNo" />
                <asp:BoundField DataField="Quantity" HeaderText="數量" SortExpression="Quantity" />
                <asp:BoundField DataField="ProductUnit" HeaderText="單位" SortExpression="ProductUnit" />
                <asp:BoundField DataField="NewPosition" HeaderText="NewPosition" SortExpression="NewPosition" Visible="False" />
                <asp:BoundField DataField="OldPosition" HeaderText="OldPosition" SortExpression="OldPosition" Visible="False" />
                <asp:BoundField DataField="InstallIn" HeaderText="InstallIn" SortExpression="InstallIn" Visible="False" />
                <asp:BoundField DataField="ComputerName" HeaderText="ComputerName" SortExpression="ComputerName" Visible="False" />
                <asp:BoundField DataField="ComputerIPv4" HeaderText="ComputerIPv4" SortExpression="ComputerIPv4" Visible="False" />
                <asp:BoundField DataField="ComputerIPv6" HeaderText="ComputerIPv6" SortExpression="ComputerIPv6" Visible="False" />
                <asp:BoundField DataField="ComputerMACAdd" HeaderText="ComputerMACAdd" SortExpression="ComputerMACAdd" Visible="False" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" Visible="False" />
            </Columns>
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
        <asp:SqlDataSource runat="server" ID="dsAssetList" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
            SelectCommand="select AssetNo, AssetType, Systematics, Brand, OriModeNumber, ProductName, OriSerialNumber, 
                                  convert(varchar, BuyDate, 111) BuyDate, Warranty, Supplier, convert(varchar, WarrantyDate, 111) WarrantyDate,
                                  OldAssetNo, ProductUnit, Quantity, NewPosition, OldPosition, InstallIn, Remark, ComputerName,
                                  ComputerIPv4, ComputerIPv6, ComputerMACAdd, 
                                  (select ClassTxt from DBDICB where FKey = '電腦資產管理    fmIAAsset       AssetType' and ClassNo = a.AssetType) AssetType_C,
                                  (select ClassTxt from DBDICB where FKey = '電腦資產管理    fmIAAsset       Systematics' and ClassNo = a.Systematics) Systematics_C,
                                  (select ClassTxt from DBDICB where FKey = '電腦資產管理    fmIAAsset       Brand' and ClassNo = a.Brand) Brand_C,
                                  (select [Name] from [Custom] where Code = a.Supplier) Supplier_C
                             from IAAssetMain a
                            where isnull(AssetNo,'') = ''"></asp:SqlDataSource>
    </asp:Panel>
    <asp:Panel runat="server" ID="IAAssetMain" GroupingText="資訊資產基本資料" Visible="true">
        <table class="TableSetting">
            <tr>
                <td class="ColWidth2 ColHeight ColBorder a3">資產編號：
                        <asp:TextBox ID="eAssetNo" runat="server" ReadOnly="True" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">資產類別：
                        <asp:DropDownList ID="eAssetType" runat="server" AutoPostBack="True" OnSelectedIndexChanged="eAssetType_SelectedIndexChanged">
                            <asp:ListItem Value="S" Text="軟體" Selected="True" />
                            <asp:ListItem Value="H" Text="硬體" />
                            <asp:ListItem Value="C" Text="耗材" />
                        </asp:DropDownList>
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">資產品項：
                        <asp:DropDownList ID="eSystematics" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">廠牌：
                        <asp:DropDownList ID="eBrand" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth2 ColHeight ColBorder a3">原廠型號：
                        <asp:TextBox ID="eOriModeNumber" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">原廠序號：
                    <asp:TextBox ID="eOriSerialNumber" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">購置日期：
                    <asp:TextBox ID="eBuyDate" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">保固年限：
                    <asp:TextBox ID="eWarranty" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth2 ColHeight ColBorder a3" colspan="3">品名：
                    <asp:TextBox ID="eProductName" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3" colspan="2">數量：
                    <asp:TextBox ID="eQuantity" runat="server" />
                    <asp:DropDownList ID="eProductUnit" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth2 ColHeight ColBorder a3" colspan="2">供應商：
                    <asp:TextBox ID="eSupplier" runat="server" />
                    <asp:TextBox ID="eSupplierName" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">保固到期日：
                    <asp:TextBox ID="eWarrantyDate" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">舊資產編號：
                    <asp:TextBox ID="eOldAssetNo" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth2 ColHeight ColBorder a3">電腦名稱：
                    <asp:TextBox ID="eComputerName" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">IPv4 位址：
                    <asp:TextBox ID="eComputerIPv4" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">IPv6 位址：
                    <asp:TextBox ID="eComputerIPv6" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">MAC 位址：
                    <asp:TextBox ID="eComputerMACAdd" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth2 ColHeight ColBorder a3">所在位置：
                    <asp:TextBox ID="eNewPosition" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3">原在位置：
                    <asp:TextBox ID="eOldPosition" runat="server" />
                </td>
                <td class="ColWidth2 ColHeight ColBorder a3" colspan="2">安裝目標電腦資產編號：
                    <asp:TextBox ID="eInstallIn" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder a3" colspan="4">
                    <asp:Label ID="lbRemark" Text="備註：" runat="server" />
                    <asp:TextBox ID="eRemark" runat="server" TextMode="MultiLine" Width="99%" Height="50px" />
                </td>
            </tr>
        </table>
        <asp:Button ID="bbCreate" runat="server" CssClass="a1" Enabled="true" Text="新增" OnClick="bbCreate_Click" />
        <asp:Button ID="bbModify" runat="server" CssClass="b2" Enabled="true" Text="修改" OnClick="bbModify_Click" />
        <asp:Button ID="bbDel" runat="server" CssClass="b4" Enabled="true" Text="刪除" OnClick="bbDel_Click" />
        <asp:Button ID="bbOK" runat="server" CssClass="b3" Enabled="false" Text="確定" OnClick="bbOK_Click" />
        <asp:Button ID="bbCancel" runat="server" CssClass="b4" Enabled="false" Text="取消" />
        <asp:Button ID="bbCheckIO" runat="server" CssClass="b1" Enabled="true" Text="進出查詢" />
        <asp:SqlDataSource ID="dsAssetData" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
            SelectCommand="select AssetNo, AssetType, Systematics, Brand, OriModeNumber, ProductName, OriSerialNumber, convert(varchar, BuyDate, 111) BuyDate, Warranty, Supplier,
                                                 convert(varchar, WarrantyDate, 111) WarrantyDate, OldAssetNo, ProductUnit, Quantity, NewPosition, OldPosition, InstallIn, Remark, ComputerName, 
                                                 ComputerIPv4, ComputerIPv6, ComputerMACAdd,
                                                 (select ClassTxt from DBDICB where FKey = '電腦資產管理    fmIAAsset       AssetType' and ClassNo = a.AssetType) AssetType_C,
                                                 (select ClassTxt from DBDICB where FKey = '電腦資產管理    fmIAAsset       Systematics' and ClassNo = a.Systematics) Systematics_C, 
                                                 (select ClassTxt from DBDICB where FKey = '電腦資產管理    fmIAAsset       Brand' and ClassNo = a.Brand) Brand_C,
                                                 (select[Name] from[Custom] where Code = a.Supplier) Supplier_C
                                            from IAAssetMain a 
                                           where isnull(AssetNo,'') = @AssetNo">
            <SelectParameters>
                <asp:Parameter Name="AssetNo" />
            </SelectParameters>
        </asp:SqlDataSource>
    </asp:Panel>
    <asp:Panel runat="server" ID="plAssetIO" GroupingText="資訊資產進出歷史" Visible="false">
    </asp:Panel>
</asp:Content>
