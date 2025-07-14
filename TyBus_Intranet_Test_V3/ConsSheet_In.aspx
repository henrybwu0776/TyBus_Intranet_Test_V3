<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ConsSheet_In.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ConsSheet_In" %>

<asp:Content ID="ConsSheet_InForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">總務耗材進貨單</a>
    </div>
    <br />
    <div>
        <asp:Label ID="eErrorMSG_Main" runat="server" CssClass="errorMessageText" Visible="false" Width="100%" />
    </div>
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuDate_Search" runat="server" CssClass="text-Right-Blue" Text="進貨日期" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eBuDateS_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eBuDateE_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbSupNo_Search" runat="server" CssClass="text-Right-Blue" Text="供應廠商" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eSupNo_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbConsName_Search" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eConsName_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col" colspan="4" />
                <td class="ColWidth-8Col" colspan="4">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="120px" />
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" OnClick="bbExit_Click" Width="120px" />
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
    <asp:Panel ID="plShowDataA" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridConsSheetA_List" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="SheetNo" DataSourceID="sdsConsSheetA_List" ForeColor="#333333" GridLines="None" PageSize="5">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="SheetNo" HeaderText="進貨單號" ReadOnly="True" SortExpression="SheetNo" />
                <asp:BoundField DataField="BuDate" DataFormatString="{0:d}" HeaderText="進貨日期" SortExpression="BuDate" />
                <asp:BoundField DataField="BuMan_C" HeaderText="建檔人" SortExpression="BuMan_C" />
                <asp:BoundField DataField="SupName" HeaderText="供應商" SortExpression="SupName" />
                <asp:BoundField DataField="SheetStatus_C" HeaderText="單據狀態" SortExpression="SheetStatus_C" />
            </Columns>
            <EditRowStyle BackColor="#7C6F57" />
            <FooterStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#666666" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#E3EAEB" />
            <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#F8FAFA" />
            <SortedAscendingHeaderStyle BackColor="#246B61" />
            <SortedDescendingCellStyle BackColor="#D4DFE1" />
            <SortedDescendingHeaderStyle BackColor="#15524A" />
        </asp:GridView>
        <asp:FormView ID="fvConsSheetA_Detail" runat="server" Width="100%" DataKeyNames="SheetNo" DataSourceID="sdsConsSheetA_Detail" OnDataBound="fvConsSheetA_Detail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOKA_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKA_Edit_Click" Text="確定" Width="120px" />
                <asp:Button ID="bbCancelA_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetNoA_Edit" runat="server" CssClass="text-Right-Blue" Text="單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetNoA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetMode_Edit" runat="server" CssClass="text-Right-Blue" Text="單據種類" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetMode_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetMode_C") %>' Width="95%" />
                            <asp:Label ID="eSheetMode_Edit" runat="server" Text='<%# Eval("SheetMode") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="處理進度" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eStatusDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="45%" />
                            <asp:Label ID="eSheetStstus_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="45%" />
                            <asp:Label ID="eSheetStatus_Edit" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSupName_Edit" runat="server" CssClass="text-Right-Blue" Text="供應廠商" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eSupNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupNo") %>' Width="20%" />
                            <asp:TextBox ID="eSupName_Edit" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("SupName") %>' Width="70%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="需求單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                            <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetNote_Edit" runat="server" CssClass="text-Right-Blue" Text="主旨" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="5">
                            <asp:TextBox ID="eSheetNote_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNote") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColBorder ColWidth-10Col MultiLine_High">
                            <asp:Label ID="lbRemarkA_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                            <asp:TextBox ID="eRemarkA_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuManA_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuMan_CA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateA_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuDateA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManA_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyMan_CA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateA_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOKA_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKA_INS_Click" Text="確定" Width="120px" />
                <asp:Button ID="bbCancelA_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetNoA_INS" runat="server" CssClass="text-Right-Blue" Text="單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetNoA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetMode_INS" runat="server" CssClass="text-Right-Blue" Text="單據種類" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetMode_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetMode_C") %>' Width="95%" />
                            <asp:Label ID="eSheetMode_INS" runat="server" Text='<%# Eval("SheetMode") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetStatus_INS" runat="server" CssClass="text-Right-Blue" Text="處理進度" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eStatusDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="45%" />
                            <asp:Label ID="eSheetStstus_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="45%" />
                            <asp:Label ID="eSheetStatus_INS" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSupName_INS" runat="server" CssClass="text-Right-Blue" Text="供應廠商" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eSupNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupNo") %>' Width="20%" />
                            <asp:TextBox ID="eSupName_INS" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("SupName") %>' Width="70%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="需求單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                            <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetNote_INS" runat="server" CssClass="text-Right-Blue" Text="主旨" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="5">
                            <asp:TextBox ID="eSheetNote_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNote") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColBorder ColWidth-10Col MultiLine_High">
                            <asp:Label ID="lbRemarkA_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                            <asp:TextBox ID="eRemarkA_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuManA_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuMan_CA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateA_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuDateA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManA_INS" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyMan_CA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateA_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNewA_Empty" runat="server" CssClass="button-Blue" CommandName="New" Text="新增" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNewA_List" runat="server" CssClass="button-Blue" CausesValidation="true" CommandName="New" Text="新增" Width="120px" />
                <asp:Button ID="bbEditA_List" runat="server" CssClass="button-Black" CausesValidation="false" CommandName="Edit" Text="修改" Width="120px" />
                <asp:Button ID="bbInStoreA_List" runat="server" CssClass="button-Blue" CausesValidation="false" OnClick="bbInStoreA_List_Click" Text="整單入庫" Width="120px" />
                <asp:Button ID="bbAbortA_List" runat="server" CssClass="button-Black" CausesValidation="false" OnClick="bbAbortA_List_Click" Text="整單作廢" Width="120px" />
                <asp:Button ID="bbDeleteA_List" runat="server" CssClass="button-Red" CausesValidation="false" OnClick="bbDeleteA_List_Click" Text="刪除" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetNoA_List" runat="server" CssClass="text-Right-Blue" Text="單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetNoA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetMode_List" runat="server" CssClass="text-Right-Blue" Text="單據種類" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetMode_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetMode_C") %>' Width="95%" />
                            <asp:Label ID="eSheetMode_List" runat="server" Text='<%# Eval("SheetMode") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetStatus_List" runat="server" CssClass="text-Right-Blue" Text="處理進度" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eStatusDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="45%" />
                            <asp:Label ID="eSheetStstus_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="45%" />
                            <asp:Label ID="eSheetStatus_List" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSupName_List" runat="server" CssClass="text-Right-Blue" Text="供應廠商" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eSupNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupNo") %>' Width="20%" />
                            <asp:Label ID="eSupName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupName") %>' Width="70%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="需求單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetNote_List" runat="server" CssClass="text-Right-Blue" Text="主旨" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="5">
                            <asp:Label ID="eSheetNote_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNote") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColBorder ColWidth-10Col MultiLine_High">
                            <asp:Label ID="lbRemarkA_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                            <asp:TextBox ID="eRemarkA_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("RemarkA") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuManA_List" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuMan_CA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateA_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuDateA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManA_List" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyMan_CA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateA_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
            </ItemTemplate>
        </asp:FormView>
        <asp:SqlDataSource ID="sdsConsSheetA_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select a.SheetNo, a.BuDate, e.[Name] BuMan_C, c.[Name] SupName, a.TotalAmount, d.ClassTxt SheetStatus_C
  from ConsSheetA a left join ConsSheetB b on b.SheetNo = a.SheetNo 
                    left join [Custom] c on c.Code = a.SupNo and c.[Types] = 'S' 
                    left join Consumables c2 on c2.ConsNo = b.ConsNo 
                    left join DBDICB d on d.ClassNo = a.SheetStatus and d.FKey = '總務耗材進出單  fmConsSheetA    SheetStatus'
                    left join Employee e on e.EmpNo = a.BuMan 
 where isnull(a.SheetMode, '') = ''"></asp:SqlDataSource>
        <asp:SqlDataSource ID="sdsConsSheetA_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select a.SheetNo, a.SheetMode, d.ClassTxt SheetMode_C, a.SupNo, c.[Name] SupName, a.BuDate, a.BuMan, e.[Name] BuMan_C, 
       a.SheetStatus, d2.ClassTxt SheetStatus_C, a.StatusDate, a.RemarkA, a.ModifyMan, e2.[Name] ModifyMan_C, a.ModifyDate, 
	   a.SheetNote, a.DepNo, f.[Name] DepName 
  from ConsSheetA a left  join DBDICB d on d.ClassNo = a.SheetMode and d.FKey = '總務課耗材進出單ConsSheetA      SheetMode'
                    left join DBDICB d2 on d2.ClassNo = a.SheetStatus and d2.FKey = '總務耗材進出單  fmConsSheetA    SheetStatus'
                    left join Employee e on e.EmpNo = a.BuMan
                    left join Employee e2 on e2.EmpNo = a.ModifyMan
                    left join [Custom] c on c.Code = a.SupNo and c.[Types] = 'S'
					left join Department f on f.DepNo = a.DepNo
 where SheetNo = @SheetNo">
            <SelectParameters>
                <asp:ControlParameter ControlID="gridConsSheetA_List" Name="SheetNo" PropertyName="SelectedValue" />
            </SelectParameters>
        </asp:SqlDataSource>
    </asp:Panel>
    <asp:Panel ID="plShowDataB" runat="server" CssClass="ShowPanel-Detail">
        <asp:GridView ID="gridConsSheetB_List" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" DataKeyNames="SheetNoItems" DataSourceID="sdsConsSheetB_List" PageSize="5">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" ItemStyle-Width="60px" ItemStyle-CssClass="fixedWidth" />
                <asp:BoundField DataField="SheetNoItems" HeaderText="SheetNoItems" ReadOnly="True" SortExpression="SheetNoItems" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" ItemStyle-Width="60px" ItemStyle-CssClass="fixedWidth" />
                <asp:BoundField DataField="ConsName" HeaderText="品名" SortExpression="ConsName" />
                <asp:BoundField DataField="Price" HeaderText="Price" SortExpression="Price" Visible="False" />
                <asp:BoundField DataField="Quantity" HeaderText="數量" SortExpression="Quantity" ItemStyle-Width="100px" ItemStyle-CssClass="fixedWidth" />
                <asp:BoundField DataField="ConsUnit_C" HeaderText="單位" SortExpression="ConsUnit_C" ItemStyle-Width="60px" ItemStyle-CssClass="fixedWidth" />
            </Columns>
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="#FFFFCC" />
            <PagerStyle BackColor="#FFFFCC" ForeColor="#330099" HorizontalAlign="Center" />
            <RowStyle BackColor="White" ForeColor="#330099" />
            <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="#663399" />
            <SortedAscendingCellStyle BackColor="#FEFCEB" />
            <SortedAscendingHeaderStyle BackColor="#AF0101" />
            <SortedDescendingCellStyle BackColor="#F6F0C0" />
            <SortedDescendingHeaderStyle BackColor="#7E0000" />
        </asp:GridView>
        <asp:FormView ID="fvConsSheetB_Detail" runat="server" DataKeyNames="SheetNoItems" DataSourceID="sdsConsSheetB_Detail" OnDataBound="fvConsSheetB_Detail_DataBound" Width="100%">
            <EditItemTemplate>
                <asp:Button ID="bbOKB_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKB_Edit_Click" Text="確定" />
                <asp:Button ID="bbCancelB_Edit" runat="server" CausesValidation="False" CommandName="Cancel" Text="取消" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetNoItems_Edit" runat="server" CssClass="text-Right-Blue" Text="項次" Width="95%" />
                            <asp:Label ID="eSheetNoB_Edit" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoItems_Edit" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbConsNo_Edit" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="7">
                            <asp:TextBox ID="eConsNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="25%" />
                            <asp:TextBox ID="eConsName_Edit" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("ConsName") %>' Width="70%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItemStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="處理狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItemStatus_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                            <asp:Label ID="eItemStatus_Edit" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="2">
                            <asp:Label ID="lbRemarkB_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="2" colspan="7">
                            <asp:TextBox ID="eRemarkB_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbQuantity_Edit" runat="server" CssClass="text-Right-Blue" Text="數量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eQuantity_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Quantity") %>' Width="50%" />
                            <asp:DropdownList ID="ddlConsUnit_C_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlConsUnit_C_Edit_SelectedIndexChanged" Width="40%" />
                            <asp:Label ID="eConsUnit_Edit" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateB_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuDateB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuManB_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuMan_CB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateB_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManB_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyMan_CB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="60%" />
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
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOKB_INS" runat="server" CausesValidation="True" OnClick="bbOKB_INS_Click" Text="確定" />
                <asp:Button ID="bbCancelB_INS" runat="server" CausesValidation="False" CommandName="Cancel" Text="取消" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetNoItems_INS" runat="server" CssClass="text-Right-Blue" Text="項次" Width="95%" />
                            <asp:Label ID="eSheetNoB_INS" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItems_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoItems_INS" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbConsNo_INS" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="7">
                            <asp:TextBox ID="eConsNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="25%" />
                            <asp:TextBox ID="eConsName_INS" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("ConsName") %>' Width="70%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItemStatus_INS" runat="server" CssClass="text-Right-Blue" Text="處理狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItemStatus_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                            <asp:Label ID="eItemStatus_INS" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="2">
                            <asp:Label ID="lbRemarkB_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="2" colspan="7">
                            <asp:TextBox ID="eRemarkB_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbQuantity_INS" runat="server" CssClass="text-Right-Blue" Text="數量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eQuantity_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Quantity") %>' Width="50%" />
                            <asp:DropdownList ID="ddlConsUnit_C_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlConsUnit_C_INS_SelectedIndexChanged" Width="40%" />
                            <asp:Label ID="eConsUnit_INS" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateB_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuDateB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuManB_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuMan_CB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateB_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManB_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyMan_CB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="60%" />
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
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNewB_Empty" runat="server" CssClass="button-Blue" CommandName="New" Text="新增明細" Width="120px" />
                <asp:Button ID="bbPickup_Empty" runat="server" CssClass="button-Black" OnClick="bbPickup_Empty_Click" Text="明細挑單" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNewB_List" runat="server" CssClass="button-Black" CommandName="New" Text="新增明細" Width="120px" />
                <asp:Button ID="bbPickup_List" runat="server" CssClass="button-Blue" OnClick="bbPickup_List_Click" Text="明細挑單" Width="120px" />
                <asp:Button ID="bbEditB_List" runat="server" CssClass="button-Black" CommandName="Edit" Text="修改明細" Width="120px" />
                <asp:Button ID="bbInstoreB_List" runat="server" CssClass="button-Blue" OnClick="bbInstoreB_List_Click" Text="明細入庫" Width="120px" />
                <asp:Button ID="bbAbortB_List" runat="server" CssClass="button-Black" OnClick="bbAbortB_List_Click" Text="明細作廢" Width="120px" />
                <asp:Button ID="bbDeleteB_List" runat="server" CssClass="button-Red" OnClick="bbDeleteB_List_Click" Text="刪除明細" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetNoItems_List" runat="server" CssClass="text-Right-Blue" Text="項次" Width="95%" />
                            <asp:Label ID="eSheetNoB_List" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoItems_List" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbConsNo_List" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="7">
                            <asp:Label ID="eConsNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="25%" />
                            <asp:Label ID="eConsName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="70%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItemStatus_List" runat="server" CssClass="text-Right-Blue" Text="處理狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItemStatus_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                            <asp:Label ID="eItemStatus_List" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="2">
                            <asp:Label ID="lbRemarkB_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="2" colspan="7">
                            <asp:TextBox ID="eRemarkB_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("RemarkB") %>' Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbQuantity_List" runat="server" CssClass="text-Right-Blue" Text="數量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eQuantity_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Quantity") %>' Width="50%" />
                            <asp:Label ID="eConsUnit_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="40%" />
                            <asp:Label ID="eConsUnit_List" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateB_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuDateB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuManB_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuMan_CB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateB_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManB_List" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyMan_CB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="60%" />
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
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:Panel ID="plShowData" runat="server">
    </asp:Panel>
    <asp:SqlDataSource ID="sdsConsSheetB_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select b.SheetNoItems, b.Items, c.ConsName, b.Price, b.Quantity, d.ClassTxt ConsUnit_C from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo left join DBDICB d on d.ClassNo = b.ConsUnit and d.FKey = '耗材庫存        CONSUMABLES     ConsUnit' where b.SheetNo = @SheetNo ">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridConsSheetA_List" Name="SheetNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsConsSheetB_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select b.SheetNoItems, b.SheetNo, b.Items, b.ConsNo, c.ConsName, b.Quantity, b.RemarkB, b.ItemStatus, d2.ClassTxt ItemStatus_C, 
       b.BuMan, e.[Name] BuMan_C, b.BuDate, b.ModifyMan, e2.[Name] ModifyMan_C, b.ModifyDate, b.ConsUnit, d.ClassTxt ConsUnit_C
  from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo
                                       left join DBDICB d on d.ClassNo = b.ConsUnit and d.FKey = '耗材庫存        CONSUMABLES     ConsUnit'
                                       left join DBDICB d2 on d2.ClassNo = b.ItemStatus and d2.FKey = '總務課耗材進出單ConsSheetB      ItemStatus'
                                       left join employee e on e.EmpNo = b.BuMan
                                       left join Employee e2 on e2.EmpNo = b.ModifyMan
 where SheetNoItems = @SheetNoItems">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridConsSheetB_List" Name="SheetNoItems" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:Panel ID="plPickupDetail" runat="server" CssClass="ShowPanel-Detail">
        <div>
            <asp:Button ID="bbSelectAll_P" runat="server" CssClass="button-Black" OnClick="bbSelectAll_P_Click" Text="全選" Width="120px" />
            <asp:Button ID="bbUnselectAll_P" runat="server" CssClass="button-Blue" OnClick="bbUnselectAll_P_Click" Text="取消全選" Width="120px" />
            <asp:Button ID="bbOK_P" runat="server" CssClass="button-Black" OnClick="bbOK_P_Click" Text="確定" Width="120px" />
            <asp:Button ID="bbCancel_P" runat="server" CssClass="button-Red" OnClick="bbCancel_P_Click" Text="取消" Width="120px" />
        </div>
        <asp:GridView ID="gridPickup_List" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" DataKeyNames="SheetNoItems" DataSourceID="sdsPickup_List" BackColor="White" BorderColor="#336666" BorderStyle="Double" BorderWidth="3px" CellPadding="4" GridLines="Horizontal">
            <Columns>
                <asp:TemplateField ItemStyle-Width="60px" ItemStyle-CssClass="fixedWidth">
                    <ItemTemplate>
                        <asp:CheckBox ID="cbChoise" runat="server" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="SheetNoItems" HeaderText="SheetNoItems" ReadOnly="True" SortExpression="SheetNoItems" Visible="False" />
                <asp:BoundField DataField="SheetNo" HeaderText="原單號" SortExpression="SheetNo" ItemStyle-Width="100px" ItemStyle-CssClass="fixedWidth" />
                <asp:BoundField DataField="Items" HeaderText="原項次" SortExpression="Items" ItemStyle-Width="60px" ItemStyle-CssClass="fixedWidth" />
                <asp:BoundField DataField="ConsNo" HeaderText="料號" SortExpression="ConsNo" ItemStyle-Width="120px" ItemStyle-CssClass="fixedWidth" />
                <asp:BoundField DataField="ConsName" HeaderText="品名" SortExpression="ConsName" />
                <asp:BoundField DataField="Quantity" HeaderText="採購量" SortExpression="Quantity" ItemStyle-Width="60px" ItemStyle-CssClass="fixedWidth" />
                <asp:BoundField DataField="StockQty" HeaderText="庫存量" SortExpression="StockQty" ItemStyle-Width="60px" ItemStyle-CssClass="fixedWidth" />
                <asp:BoundField DataField="RemarkB" HeaderText="備註" SortExpression="RemarkB" />
            </Columns>
            <FooterStyle BackColor="White" ForeColor="#333333" />
            <HeaderStyle BackColor="#336666" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#336666" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="White" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#339966" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F7F7F7" />
            <SortedAscendingHeaderStyle BackColor="#487575" />
            <SortedDescendingCellStyle BackColor="#E5E5E5" />
            <SortedDescendingHeaderStyle BackColor="#275353" />
        </asp:GridView>
        <asp:SqlDataSource ID="sdsPickup_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select b.SheetNoItems, b.SheetNo, b.Items, b.ConsNo, c.ConsName, b.Quantity, c.StockQty, b.RemarkB
  from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo  where isnull(SheetNoItems, '') = ''"></asp:SqlDataSource>
    </asp:Panel>
</asp:Content>
