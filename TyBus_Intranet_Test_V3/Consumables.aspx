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
                <td class="ColWidth-8Col" colspan="8">
                    <asp:Label ID="eErrorMSG_Main" runat="server" CssClass="errorMessageText" Visible="false" />
                </td>
            </tr>
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
                    <asp:Label ID="lbBrand_Search" runat="server" CssClass="text-Right-Blue" Text="廠牌" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eBrand_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
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
        <asp:GridView ID="gridShowList" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#999999" BorderStyle="None" BorderWidth="1px" CellPadding="3" DataKeyNames="ConsNo" DataSourceID="sdsShowDataList" GridLines="Vertical" OnPageIndexChanging="gridShowList_PageIndexChanging" PageSize="5">
            <AlternatingRowStyle BackColor="Gainsboro" />
            <Columns>
                <asp:CommandField ButtonType="Button" InsertVisible="False" ShowCancelButton="False" ShowSelectButton="True" />
                <asp:BoundField DataField="ConsNo" HeaderText="料號" ReadOnly="True" SortExpression="ConsNo" />
                <asp:BoundField DataField="ConsName" HeaderText="品名" SortExpression="ConsName" />
                <asp:BoundField DataField="ConsType" HeaderText="耗材類別" SortExpression="ConsType" />
                <asp:BoundField DataField="StockQty" HeaderText="庫存量" SortExpression="StockQty" />
                <asp:BoundField DataField="ConsUnit" HeaderText="單位" SortExpression="ConsUnit" />
                <asp:BoundField DataField="AvgPrice" HeaderText="平均成本" SortExpression="AvgPrice" />
                <asp:BoundField DataField="StoreLocation" HeaderText="存放庫位" SortExpression="StoreLocation" />
                <asp:BoundField DataField="Brand" HeaderText="廠牌" SortExpression="Brand" />
            </Columns>
            <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
            <HeaderStyle BackColor="#000084" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <RowStyle BackColor="#EEEEEE" ForeColor="Black" />
            <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F1F1F1" />
            <SortedAscendingHeaderStyle BackColor="#0000A9" />
            <SortedDescendingCellStyle BackColor="#CAC9C9" />
            <SortedDescendingHeaderStyle BackColor="#000065" />
        </asp:GridView>
        <asp:FormView ID="fvShowList" runat="server" DataKeyNames="ConsNo" DataSourceID="sdsShowDetail" OnDataBound="fvShowList_DataBound" Width="100%">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CssClass="button-Black" Text="確定" OnClick="bbOK_Edit_Click" Width="120px" />
                <asp:Button ID="bbCancel_Edit" runat="server" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsNo_Edit" runat="server" CssClass="text-Right-Blue" Text="料號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eConsNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsName_Edit" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:TextBox ID="eConsName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsType_Edit" runat="server" CssClass="text-Right-Blue" Text="類別" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:DropDownList ID="eConsType_C_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="eConsType_C_Edit_SelectedIndexChanged" Width="95%" />
                            <asp:Label ID="eConsType_Edit" runat="server" Text='<%# Eval("ConsType") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBrand_Edit" runat="server" CssClass="text-Right-Blue" Text="廠牌" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:DropDownList ID="eBrand_C_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="eBrand_C_Edit_SelectedIndexChanged" Width="95%" />
                            <asp:Label ID="eBrand_Edit" runat="server" Text='<%# Eval("Brand") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsSpec_Edit" runat="server" CssClass="text-Right-Blue" Text="規格" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:TextBox ID="eConsSpec_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsSpec") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsSpec2_Edit" runat="server" CssClass="text-Right-Blue" Text="尺寸" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eConsSpec2_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsSpec2") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStockQTY_Edit" runat="server" CssClass="text-Right-Blue" Text="在庫數" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStockQTY_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StockQty") %>' Width="55%" />
                            <asp:DropDownList ID="eConsUnit_C_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="eConsUnit_C_Edit_SelectedIndexChanged" Width="35%" />
                            <asp:Label ID="eConsUnit_Edit" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsColor" runat="server" CssClass="text-Right-Blue" Text="顏色" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eConsColor_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsColor") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAVGPrice_Edit" runat="server" CssClass="text-Right-Blue" Text="平均單價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAVGPrice_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AVGPrice") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="cbIsStopUse_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbIsStopUse_Edit_CheckedChanged" Text="停用" />
                            <asp:Label ID="eIsStopUse_Edit" runat="server" Text='<%# Eval("IsStopUse") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="cbIsInorder_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbIsInorder_Edit_CheckedChanged" Text="採購中" />
                            <asp:Label ID="eIsInorder_Edit" runat="server" Text='<%# Eval("IsInorder") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStoreLocation_Edit" runat="server" CssClass="text-Right-Blue" Text="存放庫位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eStoreLocation_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StoreLocation") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4">
                            <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5" rowspan="4">
                            <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Remark") %>' Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLastInDate_Edit" runat="server" CssClass="text-Right-Blue" Text="最後進貨日" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLastInDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastInDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLastInPrice_Edit" runat="server" CssClass="text-Right-Blue" Text="最後進價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLastInPrice_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastInPrice") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLastOutDate_Edit" runat="server" CssClass="text-Right-Blue" Text="最後出貨日" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLastOutDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastOutDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuMan_Edit" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                            <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_Edit" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                            <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
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
            </EditItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbInsert_Empty" runat="server" CssClass="button-Blue" Text="新增" CommandName="New" Width="120px" />
            </EmptyDataTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CssClass="button-Black" CausesValidation="True" OnClick="bbOK_INS_Click" Text="確定" />
                <asp:Button ID="bbCancel_INS" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsNo_INS" runat="server" CssClass="text-Right-Blue" Text="料號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eConsNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsName_INS" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:TextBox ID="eConsName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsType_INS" runat="server" CssClass="text-Right-Blue" Text="類別" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:DropDownList ID="eConsType_C_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="eConsType_C_INS_SelectedIndexChanged" Width="95%" />
                            <asp:Label ID="eConsType_INS" runat="server" Text='<%# Eval("ConsType") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBrand_INS" runat="server" CssClass="text-Right-Blue" Text="廠牌" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:DropDownList ID="eBrand_C_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="eBrand_C_INS_SelectedIndexChanged" Width="95%" />
                            <asp:Label ID="eBrand_INS" runat="server" Text='<%# Eval("Brand") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsSpec_INS" runat="server" CssClass="text-Right-Blue" Text="規格" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:TextBox ID="eConsSpec_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsSpec") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsSpec2_INS" runat="server" CssClass="text-Right-Blue" Text="尺寸" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eConsSpec2_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsSpec2") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStockQTY_INS" runat="server" CssClass="text-Right-Blue" Text="在庫數" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eStockQTY_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StockQty") %>' Width="55%" />
                            <asp:DropDownList ID="eConsUnit_C_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="eConsUnit_C_INS_SelectedIndexChanged" Width="35%" />
                            <asp:Label ID="eConsUnit_INS" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsColor" runat="server" CssClass="text-Right-Blue" Text="顏色" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eConsColor_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsColor") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAVGPrice_INS" runat="server" CssClass="text-Right-Blue" Text="平均單價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eAVGPrice_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AVGPrice") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="cbIsStopUse_INS" runat="server" CssClass="text-Left-Black" Text="停用" />
                            <asp:Label ID="eIsStopUse_INS" runat="server" Text='<%# Eval("IsStopUse") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="cbIsInorder_INS" runat="server" CssClass="text-Left-Black" Text="採購中" />
                            <asp:Label ID="eIsInorder_INS" runat="server" Text='<%# Eval("IsInorder") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStoreLocation_INS" runat="server" CssClass="text-Right-Blue" Text="存放庫位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eStoreLocation_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StoreLocation") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4">
                            <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5" rowspan="4">
                            <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Remark") %>' Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLastInDate_INS" runat="server" CssClass="text-Right-Blue" Text="最後進貨日" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLastInDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastInDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLastInPrice_INS" runat="server" CssClass="text-Right-Blue" Text="最後進價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLastInPrice_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastInPrice") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLastOutDate_INS" runat="server" CssClass="text-Right-Blue" Text="最後出貨日" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLastOutDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastOutDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuMan_INS" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                            <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_INS" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                            <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
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
            </InsertItemTemplate>
            <ItemTemplate>
                <asp:Button ID="bbInsert_List" runat="server" CssClass="button-Black" Text="新增" CommandName="New" Width="120px" />
                <asp:Button ID="bbEdit_List" runat="server" CssClass="button-Blue" Text="修改" CommandName="Edit" Width="120px" />
                <asp:Button ID="bbDelete_List" runat="server" CssClass="button-Red" Text="刪除" OnClick="bbDelete_List_Click" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsNo_List" runat="server" CssClass="text-Right-Blue" Text="料號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eConsNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsName_List" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eConsName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsType_List" runat="server" CssClass="text-Right-Blue" Text="類別" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eConsType_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsType_C") %>' Width="95%" />
                            <asp:Label ID="eConsType_List" runat="server" Text='<%# Eval("ConsType") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBrand_List" runat="server" CssClass="text-Right-Blue" Text="廠牌" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBrand_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Brand_C") %>' Width="95%" />
                            <asp:Label ID="eBrand_List" runat="server" Text='<%# Eval("Brand") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsSpec_List" runat="server" CssClass="text-Right-Blue" Text="規格" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eConsSpec_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsSpec") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsSpec2_List" runat="server" CssClass="text-Right-Blue" Text="尺寸" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eConsSpec2_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsSpec2") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStockQTY_List" runat="server" CssClass="text-Right-Blue" Text="在庫數" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStockQTY_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StockQty") %>' Width="55%" />
                            <asp:Label ID="eConsUnit_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="35%" />
                            <asp:Label ID="eConsUnit_List" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsColor" runat="server" CssClass="text-Right-Blue" Text="顏色" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eConsColor_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsColor") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAVGPrice_List" runat="server" CssClass="text-Right-Blue" Text="平均單價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAVGPrice_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AVGPrice") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="cbIsStopUse_List" runat="server" CssClass="text-Left-Black" Text="停用" />
                            <asp:Label ID="eIsStopUse_List" runat="server" Text='<%# Eval("IsStopUse") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="cbIsInorder_List" runat="server" CssClass="text-Left-Black" Text="採購中" />
                            <asp:Label ID="eIsInorder_List" runat="server" Text='<%# Eval("IsInorder") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStoreLocation_List" runat="server" CssClass="text-Right-Blue" Text="存放庫位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStoreLocation_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StoreLocation") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5" rowspan="4">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Remark") %>' Width="97%" Height="97%" Enabled="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLastInDate_List" runat="server" CssClass="text-Right-Blue" Text="最後進貨日" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLastInDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastInDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLastInPrice_List" runat="server" CssClass="text-Right-Blue" Text="最後進價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLastInPrice_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastInPrice") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLastOutDate_List" runat="server" CssClass="text-Right-Blue" Text="最後出貨日" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLastOutDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastOutDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuMan_List" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                            <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_List" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                            <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
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
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsShowDataList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select c.ConsNo, c.ConsName, d1.ClassTxt ConsType, c.StockQty, d2.ClassTxt ConsUnit, c.AvgPrice, c.StoreLocation, c.Brand
from Consumables c left join DBDICB d1 on d1.ClassNo = c.ConsType and d1.FKey = '耗材庫存        CONSUMABLES     ConsType' 
left join DBDICB d2 on d2.ClassNo = c.ConsUnit and d2.FKey = '耗材庫存        CONSUMABLES     ConsUnit' 
 where isnull(c.ConsNo, '') = '' "></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsShowDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select c.ConsNo, c.ConsName, c.ConsType, d1.ClassTxt ConsType_C, c.ConsUnit, d2.ClassTxt ConsUnit_C, c.ConsColor, c.ConsSpec, c.ConsSpec2, 
       c.Brand, b.BrandName Brand_C, c.StockQty, c.AVGPrice, c.LastInDate, c.LastOutDate, c.StoreLocation, c.IsStopUse, c.IsInorder, c.LastInPrice, 
       c.BuDate, c.BuMan, e1.[Name] as BuManName, c.ModifyDate, c.ModifyMan, e2.[Name] as ModifyManName, c.Remark 
  from Consumables c left join Employee e1 on e1.EmpNo = c.BuMan 
                     left join Employee e2 on e2.EmpNo = c.ModifyMan
					 left join DBDICB d1 on d1.ClassNo = c.ConsType and d1.FKey = '耗材庫存        CONSUMABLES     ConsType'
					 left join DBDICB d2 on d2.ClassNo = c.ConsUnit and d2.FKey = '耗材庫存        CONSUMABLES     ConsUnit'
					 left join ConsBrand b on b.BrandCode = c.Brand and b.BelongGroup = '02'
where c.ConsNo = @ConsNo">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridShowList" Name="ConsNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:Panel ID="plReport" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCloseReport" runat="server" CssClass="button-Red" Text="結束預覽" OnClick="bbCloseReport_Click" Width="120px" />
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
</asp:Content>
