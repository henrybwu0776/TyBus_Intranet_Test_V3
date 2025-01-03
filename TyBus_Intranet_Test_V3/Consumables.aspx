<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="Consumables.aspx.cs" Inherits="TyBus_Intranet_Test_V3.Consumables" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="ConsumablesForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">總務耗材管理</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbConsNo_Search" runat="server" CssClass="text-Right-Blue" Text="料號" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eConsNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eConsNo_Search_TextChanged" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbConsName_Search" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:TextBox ID="eConsName_Search" runat="server" CssClass="text-Left-Black" Width="65%" />
                    <asp:Label ID="lbErrorMSG_ConsName" runat="server" CssClass="text-Left-Red" Text="" Width="30%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbConsType_Search" runat="server" CssClass="text-Right-Blue" Text="耗材分類" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlConsType_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBrand_Search" runat="server" CssClass="text-Right-Blue" Text="廠牌" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eBrand_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPrintCheckList" runat="server" CssClass="button-Blue" Text="產生盤點單" OnClick="bbPrintCheckList_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col" colspan="3">
                    <asp:FileUpload ID="fuExcel" runat="server" Width="65%" />
                    <asp:Button ID="bbUpdateReserve_Search" runat="server" CssClass="button-Blue" Text="匯入盤點量" OnClick="bbUpdateReserve_Search_Click" Width="30%" />
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
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gvShowList" runat="server" BackColor="White" BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px" CellPadding="4" Width="100%"
            AutoGenerateColumns="False" DataKeyNames="ConsNo" DataSourceID="dsShowList" AllowPaging="True" OnPageIndexChanging="gvShowList_PageIndexChanging"
            PageSize="5" OnSelectedIndexChanged="gvShowList_SelectedIndexChanged">
            <Columns>
                <asp:ButtonField ButtonType="Button" Text="選" />
                <asp:BoundField DataField="ConsNo" HeaderText="料號" ReadOnly="True" SortExpression="ConsNo" />
                <asp:BoundField DataField="ConsName" HeaderText="品名" SortExpression="ConsName" />
                <asp:BoundField DataField="ConsType" HeaderText="類別" SortExpression="ConsType" />
                <asp:BoundField DataField="StockQty" HeaderText="庫存量" SortExpression="StockQty" />
                <asp:BoundField DataField="ConsUnit" HeaderText="單位" SortExpression="ConsUnit" />
                <asp:BoundField DataField="AvgPrice" HeaderText="平均單價" SortExpression="AvgPrice" />
                <asp:BoundField DataField="StoreLocation" HeaderText="存放庫位" SortExpression="StoreLocation" />
                <asp:BoundField DataField="Brand" HeaderText="Brand" SortExpression="廠牌" />
            </Columns>
            <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
            <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
            <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
            <RowStyle BackColor="White" ForeColor="#003399" />
            <SelectedRowStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
            <SortedAscendingCellStyle BackColor="#EDF6F6" />
            <SortedAscendingHeaderStyle BackColor="#0D4AC4" />
            <SortedDescendingCellStyle BackColor="#D6DFDF" />
            <SortedDescendingHeaderStyle BackColor="#002876" />
        </asp:GridView>
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColWidth-8Col" colspan="8">
                    <asp:Button ID="bbInsert" runat="server" CssClass="button-Black" Text="新增" OnClick="bbInsert_Click" Width="120px" />
                    <asp:Button ID="bbEdit" runat="server" CssClass="button-Blue" Text="修改" OnClick="bbEdit_Click" Width="120px" />
                    <asp:Button ID="bbStopUse" runat="server" CssClass="button-Black" Text="停用" OnClick="bbStopUse_Click" Width="120px" />
                    <asp:Button ID="bbDelete" runat="server" CssClass="button-Red" Text="刪除" OnClick="bbDelete_Click" Width="120px" />
                    <asp:Button ID="bbOK" runat="server" CssClass="button-Blue" Text="確定" OnClick="bbOK_Click" Width="120px" />
                    <asp:Button ID="bbCancel" runat="server" CssClass="button-Red" Text="取消" OnClick="bbCancel_Click" Width="120px" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbConsNo" runat="server" CssClass="text-Right-Blue" Text="料號" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="eConsNo" runat="server" CssClass="text-Left-Black" Text='<%# vdConsNo %>' Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbConsName" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:TextBox ID="eConsName" runat="server" CssClass="text-Left-Black" Text='<%# vdConsName %>' Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbConsType" runat="server" CssClass="text-Right-Blue" Text="類別" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlConsType" runat="server" CssClass="text-Left-Black" Width="95%" />
                    <asp:Label ID="eConsType" runat="server" Text='<%# vdConsType %>' Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBrand" runat="server" CssClass="text-Right-Blue" Text="廠牌" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlBrand" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlBrand_SelectedIndexChanged" Width="95%" />
                    <asp:Label ID="eBrand" runat="server" Visible="false" Text='<%# vdBrand %>' />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbConsSpec" runat="server" CssClass="text-Right-Blue" Text="規格" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:TextBox ID="eConsSpec" runat="server" CssClass="text-Left-Black" Text='<%# vdConsSpec %>' Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbConsSpec2" runat="server" CssClass="text-Right-Blue" Text="尺寸" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eConsSpec2" runat="server" CssClass="text-Left-Black" Text='<%# vdConsSpec2 %>' Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbStockQty" runat="server" CssClass="text-Right-Blue" Text="在庫數" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="eStockQty" runat="server" CssClass="text-Left-Black" Text='<%# vdStockQty %>' Width="55%" />
                    <asp:DropDownList ID="ddlConsUnit" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="eConsUnit" runat="server" Text='<%# vdConsUnit %>' Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbConsColor" runat="server" CssClass="text-Right-Blue" Text="顏色" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eConsColor" runat="server" CssClass="text-Left-Black" Text='<%# vdConsColor %>' Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbAvgPrice" runat="server" CssClass="text-Right-Blue" Text="平均單價" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="eAvgPrice" runat="server" CssClass="text-Left-Black" Text='<%# vdAvgPrice %>' Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:CheckBox ID="cbIsStopuse" runat="server" CssClass="text-Left-Black" Text="停用" />
                    <asp:Label ID="eIsStopuse" runat="server" Visible="false" Text='<%# vdIsStopUse %>' />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:CheckBox ID="cbInOrder" runat="server" CssClass="text-Left-Black" Text="採購中" />
                    <asp:Label ID="eInOrder" runat="server" Visible="false" Text='<%# vdInOrder %>' />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbStoreLocation" runat="server" CssClass="text-Right-Blue" Text="存放庫位" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eStoreLocation" runat="server" CssClass="text-Left-Black" Text='<%# vdStoreLocation %>' Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4">
                    <asp:Label ID="lbRemark" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="5" rowspan="4">
                    <asp:TextBox ID="eRemark" runat="server" TextMode="MultiLine" Text='<%# vdRemark %>' Width="95%" Height="97%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbLastInDate" runat="server" CssClass="text-Right-Blue" Text="最後進貨日" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="eLastInDate" runat="server" CssClass="text-Left-Black" Text='<%# vdLastInDate %>' Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbLastInPrice" runat="server" CssClass="text-Right-Blue" Text="最後進價" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="eLastInPrice" runat="server" CssClass="text-Left-Black" Text='<%# vdLastInPrice %>' Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbLastOutDate" runat="server" CssClass="text-Right-Blue" Text="最後出貨日" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="eLastOutDate" runat="server" CssClass="text-Left-Black" Text='<%# vdLastOutDate %>' Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuDate" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="eBuDate" runat="server" CssClass="text-Left-Black" Text='<%# vdBuDate %>' Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuMan" runat="server" CssClass="text-Right-Blue" Text=" 建檔人" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="eBuMan" runat="server" CssClass="text-Left-Black" Text='<%# vdBuMan %>' Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbModifyDate" runat="server" CssClass="text-Right-Blue" Text="異動日" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="eModifyDate" runat="server" CssClass="text-Left-Black" Text='<%# vdModifyDate %>' Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbModifyMan" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="eModifyMan" runat="server" CssClass="text-Left-Black" Text='<%# vdModifyMan %>' Width="95%" />
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
    <asp:Panel ID="plReport" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCloseReport" runat="server" CssClass="button-Red" OnClick="bbCloseReport_Click" Text="結束預覽" Width="120px" />
        <rsweb:ReportViewer ID="rvPrint" runat="server" BackColor="" ClientIDMode="AutoID" HighlightBackgroundColor=""
            InternalBorderColor="204, 204, 204" InternalBorderStyle="Solid" InternalBorderWidth="1px"
            LinkActiveColor="" LinkActiveHoverColor="" LinkDisabledColor=""
            PrimaryButtonBackgroundColor="" PrimaryButtonForegroundColor="" PrimaryButtonHoverBackgroundColor="" PrimaryButtonHoverForegroundColor=""
            SecondaryButtonBackgroundColor="" SecondaryButtonForegroundColor="" SecondaryButtonHoverBackgroundColor="" SecondaryButtonHoverForegroundColor=""
            SplitterBackColor=""
            ToolbarDividerColor="" ToolbarForegroundColor="" ToolbarForegroundDisabledColor=""
            ToolbarHoverBackgroundColor="" ToolbarHoverForegroundColor=""
            ToolBarItemBorderColor="" ToolBarItemBorderStyle="Solid" ToolBarItemBorderWidth="1px" ToolBarItemHoverBackColor=""
            ToolBarItemPressedBorderColor="51, 102, 153" ToolBarItemPressedBorderStyle="Solid" ToolBarItemPressedBorderWidth="1px"
            ToolBarItemPressedHoverBackColor="153, 187, 226" Height="600px" Width="100%" PageCountMode="Actual">
            <LocalReport ReportPath="Report\ConsumablesP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="dsShowList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select c.ConsNo, c.ConsName, d1.ClassTxt ConsType, c.StockQty, d2.ClassTxt ConsUnit, c.AvgPrice, c.StoreLocation, c.Brand
  from Consumables c left join DBDICB d1 on d1.ClassNo = c.ConsType and d1.FKey = '耗材庫存        CONSUMABLES     ConsType' 
                     left join DBDICB d2 on d2.ClassNo = c.ConsUnit and d2.FKey = '耗材庫存        CONSUMABLES     ConsUnit' 
 where isnull(c.ConsNo, '') = ''"></asp:SqlDataSource>
</asp:Content>
