<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="IAConsumables.aspx.cs" Inherits="TyBus_Intranet_Test_V3.IAConsumables" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="IAConsumablesForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">電腦課耗材管理</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbConsNo_Search" runat="server" CssClass="text-Right-Blue" Text="耗材編號" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eConsNo_S_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="~" />
                    <asp:TextBox ID="eConsNo_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbBrand_Search" runat="server" CssClass="text-Right-Blue" Text="廠牌" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:DropDownList ID="eBrand_Search" runat="server" CssClass="text-Left-Black" Width="95%" DataSourceID="dsBrand" DataTextField="ClassTxt" DataValueField="ClassNo" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbSystematics_Search" runat="server" CssClass="text-Right-Blue" Text="品項" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:DropDownList ID="eSystematics_Search" runat="server" CssClass="text-Left-Black" Width="95%" DataSourceID="dsSystematics" DataTextField="ClassTxt" DataValueField="ClassNo" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbConsName_Search" runat="server" CssClass="text-Right-Blue" Text="品名" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eConsName_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbCorrespondModel_Search" runat="server" CssClass="text-Right-Blue" Text="適用機型" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eCorrespondModel_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbConsStatus_Search" runat="server" CssClass="text-Right-Blue" Text="庫存狀況" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:DropDownList ID="eConsStatus_Search" runat="server" CssClass="text-Left-Black" Width="95%">
                        <asp:ListItem Value="S00" Text="" Selected="True" />
                        <asp:ListItem Value="S01" Text="高於安全量" />
                        <asp:ListItem Value="S02" Text="低於安全量" />
                        <asp:ListItem Value="S03" Text="已無庫存" />
                        <asp:ListItem Value="S05" Text="採購中" />
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbOK_Search" runat="server" CssClass="button-Blue" Text="查詢" OnClick="bbOK_Search_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbStoreCheckList_Search" runat="server" CssClass="button-Black" Text="產生盤點單" OnClick="bbStoreCheckList_Search_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-7Col" colspan="3">
                    <asp:FileUpload ID="fuExcel" runat="server" Width="65%" />
                    <asp:Button ID="bbUpdateReserve_Search" runat="server" CssClass="button-Blue" Text="匯入盤點量" OnClick="bbUpdateReserve_Search_Click" Width="30%" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="IAStoreList_Search" runat="server" CssClass="button-Black" Text="庫存報表" OnClick="IAStoreList_Search_Click" Width="95%" Visible="False" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbCancel_Search" runat="server" CssClass="button-Red" Text="結束" OnClick="bbCancel_Search_Click" Width="95%" />
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
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridIAConsumablesList" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4"
            DataKeyNames="ConsNo" DataSourceID="dsIAConsumablesList" ForeColor="#333333" GridLines="None" CssClass="text-Left-Black" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="ConsNo" HeaderText="耗材編號" ReadOnly="True" SortExpression="ConsNo" />
                <asp:BoundField DataField="Systematics_C" HeaderText="品項" ReadOnly="True" SortExpression="Systematics_C" />
                <asp:BoundField DataField="Brand_C" HeaderText="廠牌" ReadOnly="True" SortExpression="Brand_C" />
                <asp:BoundField DataField="ConsName" HeaderText="品名" SortExpression="ConsName" />
                <asp:BoundField DataField="Spec_Color" HeaderText="顏色" SortExpression="Spec_Color" />
                <asp:BoundField DataField="OriModelNo" HeaderText="原廠型號" SortExpression="OriModelNo" Visible="False" />
                <asp:BoundField DataField="Quantity" HeaderText="庫存量" SortExpression="Quantity" />
                <asp:BoundField DataField="Unit_C" HeaderText="單位" ReadOnly="True" SortExpression="Unit_C" />
                <asp:BoundField DataField="SaveQty" HeaderText="安全量" SortExpression="SaveQty" Visible="False" />
                <asp:BoundField DataField="Position" HeaderText="存放庫位" SortExpression="Position" Visible="False" />
                <asp:BoundField DataField="Spec_Other" HeaderText="規格" SortExpression="Spec_Other" Visible="False" />
                <asp:BoundField DataField="CorrespondModel" HeaderText="適用機型" SortExpression="CorrespondModel" />
                <asp:BoundField DataField="LastIndate" HeaderText="最後進貨日" SortExpression="LastIndate" DataFormatString="{0:D}" />
                <asp:BoundField DataField="LastInQty" HeaderText="最後進貨量" SortExpression="LastInQty" Visible="False" />
                <asp:CheckBoxField DataField="IsInorder" HeaderText="採購中" SortExpression="IsInorder" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" />
                <asp:BoundField DataField="BuMan" HeaderText="BuMan" SortExpression="BuMan" Visible="False" />
                <asp:BoundField DataField="BuManName" HeaderText="BuManName" ReadOnly="True" SortExpression="BuManName" Visible="False" />
                <asp:BoundField DataField="BuDate" HeaderText="BuDate" SortExpression="BuDate" Visible="False" />
                <asp:BoundField DataField="ModifyMan" HeaderText="ModifyMan" SortExpression="ModifyMan" Visible="False" />
                <asp:BoundField DataField="ModifyManName" HeaderText="ModifyManName" ReadOnly="True" SortExpression="ModifyManName" Visible="False" />
                <asp:BoundField DataField="ModifyDate" HeaderText="ModifyDate" SortExpression="ModifyDate" Visible="False" />
            </Columns>
            <FooterStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#FFCC66" ForeColor="#333333" HorizontalAlign="Center" />
            <RowStyle BackColor="#FFFBD6" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="Navy" />
            <SortedAscendingCellStyle BackColor="#FDF5AC" />
            <SortedAscendingHeaderStyle BackColor="#4D0000" />
            <SortedDescendingCellStyle BackColor="#FCF6C0" />
            <SortedDescendingHeaderStyle BackColor="#820000" />
        </asp:GridView>
        <asp:FormView ID="fvIAConsumablesMamager" runat="server" DataKeyNames="ConsNo" DataSourceID="dsIAConsumablesDetail" OnDataBound="fvIAConsumablesMamager_DataBound" Width="100%">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" OnClick="bbOK_Edit_Click" CssClass="button-Blue" Text="確定" Width="120px" />
                &nbsp;
                <asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CommandName="Cancel" CssClass="button-Red" Text="取消" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbConsNo_Edit" runat="server" CssClass="text-Right-Blue" Text="耗材編號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eConsNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbConsName_Edit" runat="server" CssClass="text-Right-Blue" Text="品名" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:TextBox ID="eConsName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbSystematics_Edit" runat="server" CssClass="text-Right-Blue" Text="品項" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eSystematics_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Systematics_C") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbBrand_Edit" runat="server" CssClass="text-Right-Blue" Text="廠牌" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eBrand_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Brand") %>' Width="25%" />
                                    <asp:Label ID="eBrand_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Brand_C") %>' Width="70%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbOriModelNo_Edit" runat="server" CssClass="text-Right-Blue" Text="原廠型號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:TextBox ID="eOriModelNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OriModelNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbCorrespondModel_Edit" runat="server" CssClass="text-Right-Blue" Text="適用機型" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:TextBox ID="eCorrespondModel_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CorrespondModel") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbSpec_Color_Edit" runat="server" CssClass="text-Right-Blue" Text="顏色" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:TextBox ID="eSpec_Color_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Spec_Color") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbSpec_Other_Edit" runat="server" CssClass="text-Right-Blue" Text="規格" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:TextBox ID="eSpec_Other_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Spec_Other") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" rowspan="2">
                                    <asp:Label ID="lbLastIn_Edit" runat="server" CssClass="text-Right-Blue" Text="最後進貨" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eLastIndate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastIndate","{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbPosition_Edit" runat="server" CssClass="text-Right-Blue" Text="存放地點" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:TextBox ID="ePosition_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Position") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbPrice_Edit" runat="server" CssClass="text-Right-Blue" Text="單價" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:TextBox ID="ePrice_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Price") %>' Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eLastInQty_Edit" runat="server" CssClass="text-Right-Black" Text='<%# Eval("LastInQty") %>' Width="55%" />
                                    <asp:Label ID="eUnit_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Unit_C") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col" rowspan="2">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="4" rowspan="2">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Remark") %>' TextMode="MultiLine" Height="95%" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbQuantity_Edit" runat="server" CssClass="text-Right-Blue" Text="現有庫存量" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eQuantity_Edit" runat="server" CssClass="text-Right-Black" Text='<%# Eval("Quantity") %>' Width="35%" />
                                    <asp:DropDownList ID="ddlUnit_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="True"
                                        DataSourceID="dsUnit_Edit" DataTextField="ClassTxt" DataValueField="ClassNo" OnSelectedIndexChanged="ddlUnit_Edit_SelectedIndexChanged" Width="60%">
                                    </asp:DropDownList>
                                    <asp:Label ID="eUnit_Edit" runat="server" Text='<%# Eval("Unit") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbSaveQty_Edit" runat="server" CssClass="text-Right-Blue" Text="安全存量" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eSaveQty_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SaveQty") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-7Col">
                                    <asp:CheckBox ID="cbIsStopuse_Edit" runat="server" CssClass="text-Left-Red" Text="已停用" AutoPostBack="true" OnCheckedChanged="cbIsStopuse_Edit_CheckedChanged" />
                                    <asp:Label ID="eIsStopuse_Edit" runat="server" Text='<%# Eval("IsStopuse") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColWidth-7Col">
                                    <asp:CheckBox ID="cbIsInorder_Edit" runat="server" CssClass="text-Left-Blue" Text="採購中" Enabled="false" />
                                    <asp:Label ID="eIsInorder_Edit" runat="server" Text='<%# Eval("IsInorder") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-7Col" colspan="2" />
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" OnClick="bbOK_INS_Click" CssClass="button-Blue" Text="確定" Width="120px" />
                &nbsp;
                <asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CommandName="Cancel" CssClass="button-Red" Text="取消" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbConsNo_INS" runat="server" CssClass="text-Right-Blue" Text="耗材編號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eConsNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbConsName_INS" runat="server" CssClass="text-Right-Blue" Text="品名" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:TextBox ID="eConsName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbSystematics_INS" runat="server" CssClass="text-Right-Blue" Text="品項" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:DropDownList ID="ddlSystematics_INS" runat="server" CssClass="text-Left-Black" Width="95%" AutoPostBack="True"
                                        DataSourceID="dsSystematics_INS" DataTextField="ClassTxt" DataValueField="ClassNo" OnSelectedIndexChanged="ddlSystematics_INS_SelectedIndexChanged" />
                                    <asp:Label ID="eSystematics_INS" runat="server" Text='<%# Eval("Systematics") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbBrand_INS" runat="server" CssClass="text-Right-Blue" Text="廠牌" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:DropDownList ID="ddlBrand_INS" runat="server" CssClass="text-Left-Black" Width="95%" AutoPostBack="True"
                                        DataSourceID="dsBrand_INS" DataTextField="ClassTxt" DataValueField="ClassNo" OnSelectedIndexChanged="ddlBrand_INS_SelectedIndexChanged" />
                                    <asp:Label ID="eBrand_INS" runat="server" Text='<%# Eval("Brand") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbOriModelNo_INS" runat="server" CssClass="text-Right-Blue" Text="原廠型號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:TextBox ID="eOriModelNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OriModelNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbCorrespondModel_INS" runat="server" CssClass="text-Right-Blue" Text="適用機型" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:TextBox ID="eCorrespondModel_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CorrespondModel") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbSpec_Color_INS" runat="server" CssClass="text-Right-Blue" Text="顏色" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:TextBox ID="eSpec_Color_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Spec_Color") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbSpec_Other_INS" runat="server" CssClass="text-Right-Blue" Text="規格" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:TextBox ID="eSpec_Other_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Spec_Other") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" rowspan="2">
                                    <asp:Label ID="lbLastIn_INS" runat="server" CssClass="text-Right-Blue" Text="最後進貨" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eLastIndate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastIndate","{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbPosition_INS" runat="server" CssClass="text-Right-Blue" Text="存放地點" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:TextBox ID="ePosition_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Position") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbPrice_INS" runat="server" CssClass="text-Right-Blue" Text="單價" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:TextBox ID="ePrice_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Price") %>' Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eLastInQty_INS" runat="server" CssClass="text-Right-Black" Text='<%# Eval("LastInQty") %>' Width="55%" />
                                    <asp:Label ID="eUnit_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Unit_C") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col" rowspan="2">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="4" rowspan="2">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Remark") %>' TextMode="MultiLine" Height="95%" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbQuantity_INS" runat="server" CssClass="text-Right-Blue" Text="現有庫存量" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eQuantity_INS" runat="server" CssClass="text-Right-Black" Text='<%# Eval("Quantity") %>' Width="35%" />
                                    <asp:DropDownList ID="ddlUnit_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="True"
                                        DataSourceID="dsUnit_INS" DataTextField="ClassTxt" DataValueField="ClassNo" OnSelectedIndexChanged="ddlUnit_INS_SelectedIndexChanged" Width="60%">
                                    </asp:DropDownList>
                                    <asp:Label ID="eUnit_INS" runat="server" Text='<%# Eval("Unit") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbSaveQty_INS" runat="server" CssClass="text-Right-Blue" Text="安全存量" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eSaveQty_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SaveQty") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-7Col">
                                    <asp:CheckBox ID="cbIsStopuse_INS" runat="server" CssClass="text-Left-Red" Text="已停用" Enabled="false" />
                                    <asp:Label ID="eIsStopuse_INS" runat="server" Text='<%# Eval("IsStopuse") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColWidth-7Col">
                                    <asp:CheckBox ID="cbIsInorder_INS" runat="server" CssClass="text-Left-Blue" Text="採購中" Enabled="false" />
                                    <asp:Label ID="eIsInorder_INS" runat="server" Text='<%# Eval("IsInorder") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-7Col" colspan="2" />
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CausesValidation="False" CommandName="New" Text="新增" CssClass="button-Black" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CommandName="New" Text="新增" CssClass="button-Black" Width="120px" />
                &nbsp;
                <asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CommandName="Edit" Text="編輯" CssClass="button-Blue" Width="120px" />
                &nbsp;
                <asp:Button ID="bbDelete_List" runat="server" CausesValidation="False" OnClick="bbDelete_List_Click" Text="刪除" CssClass="button-Red" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbConsNo_List" runat="server" CssClass="text-Right-Blue" Text="耗材編號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eConsNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbConsName_List" runat="server" CssClass="text-Right-Blue" Text="品名" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:Label ID="eConsName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbSystematics_List" runat="server" CssClass="text-Right-Blue" Text="品項" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eSystematics_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Systematics_C") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbBrand_List" runat="server" CssClass="text-Right-Blue" Text="廠牌" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eBrand_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Brand") %>' Width="25%" />
                                    <asp:Label ID="eBrand_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Brand_C") %>' Width="70%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbOriModelNo_List" runat="server" CssClass="text-Right-Blue" Text="原廠型號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eOriModelNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OriModelNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbCorrespondModel_List" runat="server" CssClass="text-Right-Blue" Text="適用機型" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:Label ID="eCorrespondModel_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CorrespondModel") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbSpec_Color_List" runat="server" CssClass="text-Right-Blue" Text="顏色" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eSpec_Color_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Spec_Color") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbSpec_Other_List" runat="server" CssClass="text-Right-Blue" Text="規格" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:Label ID="eSpec_Other_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Spec_Other") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" rowspan="2">
                                    <asp:Label ID="lbLastIn_List" runat="server" CssClass="text-Right-Blue" Text="最後進貨" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eLastIndate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastIndate","{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbPosition_List" runat="server" CssClass="text-Right-Blue" Text="存放地點" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:Label ID="ePosition_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Position") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbPrice_List" runat="server" CssClass="text-Right-Blue" Text="單價" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="ePrice_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Price") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eLastInQty_List" runat="server" CssClass="text-Right-Black" Text='<%# Eval("LastInQty") %>' Width="35%" />
                                    <asp:Label ID="eUnit_C_List2" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Unit_C") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col" rowspan="2">
                                    <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="4" rowspan="2">
                                    <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Remark") %>' TextMode="MultiLine" Height="95%" Width="100%" Enabled="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbQuantity_List" runat="server" CssClass="text-Right-Blue" Text="現有庫存量" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eQuantity_List" runat="server" CssClass="text-Right-Black" Text='<%# Eval("Quantity") %>' Width="35%" />
                                    <asp:Label ID="eUnit_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Unit_C") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbSaveQty_List" runat="server" CssClass="text-Right-Blue" Text="安全存量" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eSaveQty_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SaveQty") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-7Col">
                                    <asp:CheckBox ID="cbIsStopuse_List" runat="server" CssClass="text-Left-Red" Text="已停用" />
                                    <asp:Label ID="eIsStopuse_List" runat="server" Text='<%# Eval("IsStopuse") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColWidth-7Col">
                                    <asp:CheckBox ID="cbIsInorder_List" runat="server" CssClass="text-Left-Blue" Text="採購中" Enabled="false" />
                                    <asp:Label ID="eIsInorder_List" runat="server" Text='<%# Eval("IsInorder") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-7Col" colspan="2" />
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                                    <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-7Col">
                                    <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </ItemTemplate>
        </asp:FormView>
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
            <LocalReport ReportPath="Report\IAConsumablesP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <!--廠商代碼對照資料庫-->
    <asp:SqlDataSource ID="dsBrand" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select null as ClassNo, null as ClassTxt
union all
select ClassNo, ClassTxt 
  from DBDICB
 where FKey = '電腦課耗材管理  fmIAConsumables Brand'"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsBrand_Edit" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select null as ClassNo, null as ClassTxt
union all
select ClassNo, ClassTxt 
  from DBDICB
 where FKey = '電腦課耗材管理  fmIAConsumables Brand'"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsBrand_INS" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select null as ClassNo, null as ClassTxt
union all
select ClassNo, ClassTxt 
  from DBDICB
 where FKey = '電腦課耗材管理  fmIAConsumables Brand'"></asp:SqlDataSource>
    <!--品項代碼對照資料庫-->
    <asp:SqlDataSource ID="dsSystematics" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select null as ClassNo, null as ClassTxt
union all
select ClassNo, ClassTxt 
  from DBDICB
 where FKey = '電腦課耗材管理  fmIAConsumables Systematics'"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsSystematics_Edit" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select null as ClassNo, null as ClassTxt
union all
select ClassNo, ClassTxt 
  from DBDICB
 where FKey = '電腦課耗材管理  fmIAConsumables Systematics'"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsSystematics_INS" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select null as ClassNo, null as ClassTxt
union all
select ClassNo, ClassTxt 
  from DBDICB
 where FKey = '電腦課耗材管理  fmIAConsumables Systematics'"></asp:SqlDataSource>
    <!--耗材清單 DataGrid 用-->
    <asp:SqlDataSource ID="dsIAConsumablesList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT IA.ConsNo, (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Systematics' and ClassNo = IA.Systematics) as Systematics_C, (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Brand' and ClassNo = IA.Brand) as Brand_C, ConsName, OriModelNo, Quantity, (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Unit' and ClassNo = IA.Unit) as Unit_C, SaveQty, Position, Spec_Color, Spec_Other, CorrespondModel, LastIndate, LastInQty, IsInorder, Remark, BuMan, (select [Name] from Employee where EmpNo = IA.BuMan) as BuManName, BuDate, ModifyMan, (select [Name] from Employee where EmpNo = IA.ModifyMan) as ModifyManName, ModifyDate, Price
  FROM IAConsumables as IA 
 where 1 &lt;&gt; 1"></asp:SqlDataSource>
    <!--耗材資料操作用-->
    <asp:SqlDataSource ID="dsIAConsumablesDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" DeleteCommand="DELETE FROM IAConsumables WHERE (ConsNo = @ConsNo)" InsertCommand="INSERT INTO IAConsumables(ConsNo, ConsName, Systematics, Brand, OriModelNo, CorrespondModel, Spec_Color, Spec_Other, Position, Remark, BuMan, BuDate, Price) VALUES (@ConsNo, @ConsName, @Systematics, @Brand, @OriModelNo, @CorrespondModel, @Spec_Color, @Spec_Other, @Position, @Remark, @BuMan, GETDATE(), @Price)" SelectCommand="SELECT ConsNo, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材管理  fmIAConsumables Systematics') AND (CLASSNO = IA.Systematics)) AS Systematics_C, (SELECT CLASSTXT FROM DBDICB AS DBDICB_2 WHERE (FKEY = '電腦課耗材管理  fmIAConsumables Brand') AND (CLASSNO = IA.Brand)) AS Brand_C, ConsName, OriModelNo, Quantity, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (FKEY = '電腦課耗材管理  fmIAConsumables Unit') AND (CLASSNO = IA.Unit)) AS Unit_C, SaveQty, Position, Spec_Color, Spec_Other, CorrespondModel, LastIndate, LastInQty, IsInorder, Remark, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = IA.BuMan)) AS BuManName, BuDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = IA.ModifyMan)) AS ModifyManName, ModifyDate, Systematics, Brand, Unit, IsStopuse, Price FROM IAConsumables AS IA WHERE (ConsNo = @ConsNo)"
        UpdateCommand="UPDATE IAConsumables SET ConsName = @ConsName, Brand = @Brand, OriModelNo = @OriModelNo, CorrespondModel = @CorrespondModel, Spec_Color = @Spec_Color, Spec_Other = @Spec_Other, Position = @Position, Remark = @Remark, IsStopuse = @IsStopuse, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate, Price = @Price WHERE (ConsNo = @ConsNo)">
        <DeleteParameters>
            <asp:Parameter Name="ConsNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="ConsNo" />
            <asp:Parameter Name="ConsName" />
            <asp:Parameter Name="Systematics" />
            <asp:Parameter Name="Brand" />
            <asp:Parameter Name="OriModelNo" />
            <asp:Parameter Name="CorrespondModel" />
            <asp:Parameter Name="Spec_Color" />
            <asp:Parameter Name="Spec_Other" />
            <asp:Parameter Name="Position" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="Price" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridIAConsumablesList" Name="ConsNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="ConsName" />
            <asp:Parameter Name="Brand" />
            <asp:Parameter Name="OriModelNo" />
            <asp:Parameter Name="CorrespondModel" />
            <asp:Parameter Name="Spec_Color" />
            <asp:Parameter Name="Spec_Other" />
            <asp:Parameter Name="Position" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="IsStopuse" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="Price" />
            <asp:Parameter Name="ConsNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsUnit_Edit" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材管理  fmIAConsumables Unit') ORDER BY ClassNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsUnit_INS" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材管理  fmIAConsumables Unit') ORDER BY ClassNo"></asp:SqlDataSource>
</asp:Content>
