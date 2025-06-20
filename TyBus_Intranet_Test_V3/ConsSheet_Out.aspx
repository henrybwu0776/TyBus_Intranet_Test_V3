<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ConsSheet_Out.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ConsSheet_Out" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="ConsSheet_OutForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">總務耗材請購發料單</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbBuDate_Search" runat="server" CssClass="text-Right-Blue" Text="申請日期" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eBuDate_S_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eBuDate_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Search_TextChanged" Width="95%" />
                    <br />
                    <asp:Label ID="eDepName_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbAssignMan_Search" runat="server" CssClass="text-Right-Blue" Text="申請人" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eAssignMan_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignMan_Search_TextChanged" Width="95%" />
                    <br />
                    <asp:Label ID="eAssignManName_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-7Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Blue" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColWidth-7Col">
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" OnClick="bbExit_Click" Width="95%" />
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
    <div>
        <asp:Label ID="eErrorMSG_A" runat="server" CssClass="errorMessageText" Visible="false" />
    </div>
    <asp:Panel ID="plShowData_A" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridConsSheetA_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#999999" BorderStyle="None" BorderWidth="1px" CellPadding="3" DataKeyNames="SheetNo" DataSourceID="sdsConsSheetA_List" GridLines="Vertical" PageSize="5" Width="100%">
            <AlternatingRowStyle BackColor="#DCDCDC" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="SheetNo" HeaderText="出貨單號" ReadOnly="True" SortExpression="SheetNo" />
                <asp:BoundField DataField="SheetNote" HeaderText="主旨" SortExpression="SheetNote" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="申請單位" SortExpression="DepName" />
                <asp:BoundField DataField="AssignMan" HeaderText="AssignMan" SortExpression="AssignMan" Visible="False" />
                <asp:BoundField DataField="AssignManName" HeaderText="申請人" SortExpression="AssignManName" />
                <asp:BoundField DataField="BuDate" DataFormatString="{0:d}" HeaderText="出貨日期" ReadOnly="True" SortExpression="BuDate" />
                <asp:BoundField DataField="TotalAmount" HeaderText="總金額" SortExpression="TotalAmount" Visible="False" />
                <asp:BoundField DataField="SheetStatus_C" HeaderText="處理狀態" SortExpression="SheetStatus_C" />
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
        <asp:FormView ID="fvConsSheetA_Detail" runat="server" DataKeyNames="SheetNo" DataSourceID="sdsConsSheetA_Detail" Width="100%" OnDataBound="fvConsSheetA_Detail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOKA_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKA_Edit_Click" Text="確定" Width="120px" />
                <asp:Button ID="bbCancelA_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="bbSheetNoA_Edit" runat="server" CssClass="text-Right-Blue" Text="出貨單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSheetNoA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetNote_Edit" runat="server" CssClass="text-Right-Blue" Text="主旨" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5">
                            <asp:TextBox ID="eSheetNote_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNote") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDateA_Edit" runat="server" CssClass="text-Right-Blue" Text="出貨日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eBuDateA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuManA_Edit" runat="server" CssClass="text-Right-Blue" Text="出貨人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuManA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManNameA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_High ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemarkA_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemarkA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RemarkA") %>' TextMode="MultiLine" Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="單據狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSheetStatus_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                            <asp:Label ID="eSheetStatus_Edit" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStatusDate_Edit" runat="server" CssClass="text-Right-Blue" Text="狀態日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStatusDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyManA_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyManNameA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                            <asp:Label ID="eModifyManA_Edit" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDateA_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDateA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
            <InsertItemTemplate>
                <asp:Button ID="bbOKA_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKA_INS_Click" Text="確定" Width="120px" />
                <asp:Button ID="bbCancelA_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="bbSheetNoA_INS" runat="server" CssClass="text-Right-Blue" Text="出貨單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSheetNoA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetNote_INS" runat="server" CssClass="text-Right-Blue" Text="主旨" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5">
                            <asp:TextBox ID="eSheetNote_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNote") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDateA_INS" runat="server" CssClass="text-Right-Blue" Text="出貨日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eBuDateA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuManA_INS" runat="server" CssClass="text-Right-Blue" Text="出貨人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuManA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManNameA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_High ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemarkA_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemarkA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RemarkA") %>' TextMode="MultiLine" Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetStatus_INS" runat="server" CssClass="text-Right-Blue" Text="單據狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSheetStatus_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                            <asp:Label ID="eSheetStatus_INS" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStatusDate_INS" runat="server" CssClass="text-Right-Blue" Text="狀態日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStatusDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyManA_INS" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyManNameA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                            <asp:Label ID="eModifyManA_INS" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDateA_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDateA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
            <EmptyDataTemplate>
                <asp:Button ID="bbNewA_Empty" runat="server" CssClass="button-Blue" Text="新增" CommandName="New" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNewA_List" runat="server" CssClass="button-Black" Text="新增" CommandName="New" Width="120px" />
                <asp:Button ID="bbEditA_List" runat="server" CssClass="button-Blue" Text="修改" CommandName="Edit" Width="120px" />
                <asp:Button ID="bbAbortA_List" runat="server" CssClass="button-Black" OnClick="bbAbortA_List_Click" Text="作廢" Width="120px" />
                <asp:Button ID="bbPrint_List" runat="server" CssClass="button-Blue" OnClick="bbPrint_List_Click" Text="列印簽收單" Width="120px" />
                <asp:Button ID="bbCloseA_List" runat="server" CssClass="button-Black" OnClick="bbCloseA_List_Click" Text="結案" Width="120px" />
                <asp:Button ID="bbDeleteA_List" runat="server" CssClass="button-Red" OnClick="bbDeleteA_List_Click" Text="刪除" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="bbSheetNoA_List" runat="server" CssClass="text-Right-Blue" Text="出貨單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSheetNoA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetNote_List" runat="server" CssClass="text-Right-Blue" Text="主旨" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5">
                            <asp:Label ID="eSheetNote_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNote") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDateA_List" runat="server" CssClass="text-Right-Blue" Text="出貨日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDateA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuManA_List" runat="server" CssClass="text-Right-Blue" Text="出貨人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuManA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManNameA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_High ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemarkA_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemarkA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RemarkA") %>' Enabled="false" TextMode="MultiLine" Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetStatus_List" runat="server" CssClass="text-Right-Blue" Text="單據狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSheetStatus_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                            <asp:Label ID="eSheetStatus_List" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStatusDate_List" runat="server" CssClass="text-Right-Blue" Text="狀態日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStatusDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyManA_List" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyManNameA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                            <asp:Label ID="eModifyManA_List" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDateA_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDateA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
        <asp:SqlDataSource ID="sdsConsSheetA_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT a.SheetNo, a.SheetNote, a.DepNo, d.NAME AS DepName, a.AssignMan, e.NAME AS AssignManName, 
       CONVERT (varchar(10), a.BuDate, 111) AS BuDate, a.TotalAmount, f.ClassTxt AS SheetStatus_C 
  FROM ConsSheetA AS a LEFT OUTER JOIN DEPARTMENT AS d ON d.DEPNO = a.DepNo 
                       LEFT OUTER JOIN EMPLOYEE AS e ON e.EMPNO = a.AssignMan 
                       LEFT OUTER JOIN DBDICB AS f on f.ClassNo = a.SheetStatus and f.FKey = '總務耗材進出單  fmConsSheetA    SheetStatus' 
 WHERE isnull(a.SheetMode, '') = 'OS'"></asp:SqlDataSource>
        <asp:SqlDataSource ID="sdsConsSheetA_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT a.SheetNo, CONVERT (varchar(10), a.BuDate, 111) AS BuDate, a.BuMan,
	   e2.[Name] as BuManName, a.SheetStatus, a.SheetNote, z3.ClassTxt as SheetStatus_C, 
	   convert(varchar(10), a.StatusDate, 111) as StatusDate, a.RemarkA, a.ModifyMan, e3.[Name] as ModifyManName,
	   convert(varchar(10), a.ModifyDate, 111) as ModifyDate, a.DepNo, d.[Name] as DepName, a.AssignMan,
	   e.[Name] as AssignManName
  FROM ConsSheetA a LEFT OUTER JOIN DEPARTMENT AS d ON d.DEPNO = a.DepNo 
                    LEFT OUTER JOIN EMPLOYEE AS e ON e.EMPNO = a.AssignMan
                    LEFT OUTER JOIN EMPLOYEE AS e2 ON e2.EMPNO = a.BuMan
                    LEFT OUTER JOIN EMPLOYEE AS e3 ON e3.EMPNO = a.ModifyMan
                    LEFT OUTER JOIN DBDICB as z3 on z3.ClassNo = a.SheetStatus and z3.FKey = '總務耗材進出單  fmConsSheetA    SheetStatus'
 WHERE a.SheetNo = @SheetNo">
            <SelectParameters>
                <asp:ControlParameter ControlID="gridConsSheetA_List" Name="SheetNo" PropertyName="SelectedValue" />
            </SelectParameters>
        </asp:SqlDataSource>
    </asp:Panel>
    <div>
        <asp:Label ID="eErrorMSG_B" runat="server" CssClass="errorMessageText" Visible="false" />
    </div>
    <asp:Panel ID="plShowData_B" runat="server" CssClass="ShowPanel-Detail">
        <asp:GridView ID="gridConsSheetB_List" runat="server" PageSize="5" Width="100%" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px" CellPadding="4" DataSourceID="sdsConsSheetB_List" DataKeyNames="SheetNoItems">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="SheetNoItems" HeaderText="SheetNoItems" SortExpression="SheetNoItems" Visible="False" />
                <asp:BoundField DataField="SheetNo" HeaderText="SheetNo" SortExpression="SheetNo" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                <asp:BoundField DataField="ConsName" HeaderText="品名" SortExpression="ConsName" />
                <asp:BoundField DataField="OriQty" HeaderText="申請量" SortExpression="OriQty" />
                <asp:BoundField DataField="Quantity" HeaderText="撥付量" SortExpression="Quantity" />
                <asp:BoundField DataField="ConsUnit_C" HeaderText="單位" SortExpression="ConsUnit_C" />
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
        <asp:FormView ID="fvConsSheetB_Detail" runat="server" Width="100%" DataSourceID="sdsConsSheetB_Detail" OnDataBound="fvConsSheetB_Detail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOKB_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" Text="更新" OnClick="bbOKB_Edit_Click" Width="120px" />
                <asp:Button ID="bbCancelB_Edit" runat="server" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetNoItems_Edit" runat="server" CssClass="text-Right-Blue" Text="項次" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoItems_Edit" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                            <asp:Label ID="eSheetNoB_Edit" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsNo_Edit" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5">
                            <asp:TextBox ID="eConsNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="20%" />
                            <asp:TextBox ID="eConsName_Edit" runat="server" CssClass="text-Left-Black" Enabled="false" AutoPostBack="true" OnTextChanged="eConsName_Edit_TextChanged" Text='<%# Eval("ConsName") %>' Width="75%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItemStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItemStatus_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                            <asp:Label ID="eItemStatus_Edit" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbOriQty_Edit" runat="server" CssClass="text-Right-Blue" Text="申請數量" Width="95%" />
                            <asp:Label ID="eConsUnit_Edit" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eOriQty_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OriQty") %>' Width="65%" />
                            <asp:Label ID="eConsUnit_C1_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbQuantity_Edit" runat="server" CssClass="text-Right-Blue" Text="撥付數量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="eQuantity_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Quantity") %>' Width="65%" />
                            <asp:Label ID="eConsUnit_C2_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="30%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuManB_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="95%" />
                            <asp:Label ID="eBuManB_Edit" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4">
                            <asp:Label ID="lbRemarkB_Edit" runat="server" CssClass="text-Right-Blue" Text="明細備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4" colspan="5">
                            <asp:TextBox ID="eRemarkB_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDateB_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDateB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyManB_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="95%" />
                            <asp:Label ID="eModifyManB_Edit" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDateB_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDateB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
            <InsertItemTemplate>
                <asp:Button ID="bbOKB_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKB_INS_Click" Text="確定" Width="120px" />
                <asp:Button ID="bbCancelB_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetNoItems_INS" runat="server" CssClass="text-Right-Blue" Text="項次" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItems_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoItems_INS" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                            <asp:Label ID="eSheetNoB_INS" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsNo_INS" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5">
                            <asp:TextBox ID="eConsNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="20%" />
                            <asp:TextBox ID="eConsName_INS" runat="server" CssClass="text-Left-Black" Enabled="false" OnTextChanged="eConsName_INS_TextChanged" Text='<%# Eval("ConsName") %>' Width="75%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItemStatus_INS" runat="server" CssClass="text-Right-Blue" Text="狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItemStatus_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                            <asp:Label ID="eItemStatus_INS" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbOriQty_INS" runat="server" CssClass="text-Right-Blue" Text="申請數量" Width="95%" />
                            <asp:Label ID="eConsUnit_INS" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eOriQty_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OriQty") %>' Width="65%" />
                            <asp:Label ID="eConsUnit_C1_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbQuantity_INS" runat="server" CssClass="text-Right-Blue" Text="撥付數量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="eQuantity_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Quantity") %>' Width="65%" />
                            <asp:Label ID="eConsUnit_C2_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="30%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuManB_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuMan_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="95%" />
                            <asp:Label ID="eBuManB_INS" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4">
                            <asp:Label ID="lbRemarkB_INS" runat="server" CssClass="text-Right-Blue" Text="明細備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4" colspan="5">
                            <asp:TextBox ID="eRemarkB_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDateB_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDateB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyManB_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="95%" />
                            <asp:Label ID="eModifyManB_INS" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDateB_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDateB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
            <EmptyDataTemplate>
                <asp:Button ID="bbNewB_Empty" runat="server" CssClass="button-Blue" Text="新增明細" CommandName="New" Width="120px" />
                <asp:Button ID="bbPatchNew_Empty" runat="server" CssClass="button-Black" Text="批次匯入明細" OnClick="bbPatchNew_List_Click" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbPatchNew_List" runat="server" CssClass="button-Blue" Text="批次匯入明細" OnClick="bbPatchNew_List_Click" Width="120px" />
                <asp:Button ID="bbNewB_List" runat="server" CssClass="button-Black" Text="單筆新增明細" CommandName="New" Width="120px" />
                <asp:Button ID="bbEditB_List" runat="server" CssClass="button-Blue" Text="修改明細" CommandName="Edit" Width="120px" />
                <asp:Button ID="bbAbortB_List" runat="server" CssClass="button-Black" Text="明細作廢" Width="120px" />
                <asp:Button ID="bbDeleteB_List" runat="server" CssClass="button-Red" Text="刪除明細" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetNoItems_List" runat="server" CssClass="text-Right-Blue" Text="項次" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoItems_List" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                            <asp:Label ID="eSheetNoB_List" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsNo_List" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5">
                            <asp:Label ID="eConsNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="20%" />
                            <asp:Label ID="eConsName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="75%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItemStatus_List" runat="server" CssClass="text-Right-Blue" Text="狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItemStatus_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                            <asp:Label ID="eItemStatus_List" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbOriQty_List" runat="server" CssClass="text-Right-Blue" Text="申請數量" Width="95%" />
                            <asp:Label ID="eConsUnit_List" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eOriQty_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OriQty") %>' Width="65%" />
                            <asp:Label ID="eConsUnit_C1_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbQuantity_List" runat="server" CssClass="text-Right-Blue" Text="撥付數量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eQuantity_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Quantity") %>' Width="65%" />
                            <asp:Label ID="eConsUnit_C2_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="30%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuManB_List" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="95%" />
                            <asp:Label ID="eBuManB_List" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4">
                            <asp:Label ID="lbRemarkB_List" runat="server" CssClass="text-Right-Blue" Text="明細備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4" colspan="5">
                            <asp:TextBox ID="eRemarkB_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled =" false" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDateB_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDateB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyManB_List" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="95%" />
                            <asp:Label ID="eModifyManB_List" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDateB_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDateB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
    <asp:SqlDataSource ID="sdsConsSheetB_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select b.SheetNoItems, b.SheetNo, b.Items, c.ConsName, b2.Quantity as OriQty, b.Quantity, d.ClassTxt as ConsUnit_C
  from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo
                    left join DBDICB d on d.ClassNo = b.ConsUnit and d.FKey = '耗材庫存        CONSUMABLES     ConsUnit'
					left join ConsSheetB b2 on b2.SheetNoItems = b.SourceNo
 where b.SheetNo = @SheetNo">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridConsSheetA_List" Name="SheetNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsConsSheetB_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select b.SheetNoItems, b.SheetNo, b.Items, b.ConsNo, c.ConsName, b2.Quantity as OriQty, b.Quantity, b.ConsUnit, d.ClassTxt as ConsUnit_C,
       b.ItemStatus, d2.ClassTxt ItemStatus_C, b.BuMan, e.[Name] BuMan_C, b.BuDate, b.ModifyMan, e2.[Name] ModifyMan_C, b.ModifyDate, b.RemarkB
  from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo
                    left join DBDICB d on d.ClassNo = b.ConsUnit and d.FKey = '耗材庫存        CONSUMABLES     ConsUnit'
					left join DBDICB d2 on d2.ClassNo = b.ItemStatus and d2.FKey = '總務課耗材進出單ConsSheetB      ItemStatus'
					left join ConsSheetB b2 on b2.SheetNoItems = b.SourceNo
					left join Employee e on e.EmpNo = b.BuMan
					left join Employee e2 on e2.EmpNo = b.ModifyMan
 where b.SheetNoItems = @SheetNoItems">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridConsSheetB_List" Name="SheetNoItems" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:Panel ID="plShowData_C" runat="server" CssClass="ShowPanel-Detail">
        <div>
            <asp:Button ID="bbSelectAll_Order" runat="server" CssClass="button-Blue" Text="全選" OnClick="bbSelectAll_Order_Click" Width="120px" />
            <asp:Button ID="bbUnselAll_Order" runat="server" CssClass="button-Black" Text="取消全選" OnClick="bbUnselAll_Order_Click" Width="120px" />
            <asp:Button ID="bbOKC_Order" runat="server" CssClass="button-Blue" Text="確定" OnClick="bbOKC_Order_Click" Width="120px" />
            <asp:Button ID="bbCancelC_Order" runat="server" CssClass="button-Red" Text="取消" OnClick="bbCancelC_Order_Click" Width="120px" />
        </div>
        <asp:GridView ID="gridOrderItems_List" runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="SheetNoItems" DataSourceID="sdsOrderItems_List" ForeColor="#333333" PageSize="20">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:CheckBox ID="cbChoise" runat="server" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="SheetNo" HeaderText="請購單號" SortExpression="SheetNo" />
                <asp:BoundField DataField="Items" HeaderText="原項次" SortExpression="Items" />
                <asp:BoundField DataField="SheetNoItems" HeaderText="SheetNoItems" ReadOnly="True" SortExpression="SheetNoItems" Visible="False" />
                <asp:BoundField DataField="ConsNo" HeaderText="ConsNo" SortExpression="ConsNo" Visible="False" />
                <asp:BoundField DataField="ConsName" HeaderText="品名" SortExpression="ConsName" />
                <asp:BoundField DataField="Price" HeaderText="單價" SortExpression="Price" Visible="False" />
                <asp:BoundField DataField="Quantity" HeaderText="請購數量" SortExpression="Quantity" />
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
        <asp:SqlDataSource ID="sdsOrderItems_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT b.SheetNo, b.Items, b.SheetNoItems, b.ConsNo, c.ConsName, b.Price, b.Quantity FROM ConsSheetB AS b LEFT OUTER JOIN Consumables AS c ON c.ConsNo = b.ConsNo where isnull(SheetNoItems, '') = ''"></asp:SqlDataSource>
    </asp:Panel>
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCloseReport" runat="server" CssClass="button-Red" OnClick="bbCloseReport_Click" Text="結束預覽" Width="120px" />
        <rsweb:ReportViewer ID="rvPrint" runat="server" Width="100%" BackColor="" ClientIDMode="AutoID" HighlightBackgroundColor="" InternalBorderColor="204, 204, 204" InternalBorderStyle="Solid" InternalBorderWidth="1px" LinkActiveColor="" LinkActiveHoverColor="" LinkDisabledColor="" PrimaryButtonBackgroundColor="" PrimaryButtonForegroundColor="" PrimaryButtonHoverBackgroundColor="" PrimaryButtonHoverForegroundColor="" SecondaryButtonBackgroundColor="" SecondaryButtonForegroundColor="" SecondaryButtonHoverBackgroundColor="" SecondaryButtonHoverForegroundColor="" SplitterBackColor="" ToolbarDividerColor="" ToolbarForegroundColor="" ToolbarForegroundDisabledColor="" ToolbarHoverBackgroundColor="" ToolbarHoverForegroundColor="" ToolBarItemBorderColor="" ToolBarItemBorderStyle="Solid" ToolBarItemBorderWidth="1px" ToolBarItemHoverBackColor="" ToolBarItemPressedBorderColor="51, 102, 153" ToolBarItemPressedBorderStyle="Solid" ToolBarItemPressedBorderWidth="1px" ToolBarItemPressedHoverBackColor="153, 187, 226">
            <LocalReport ReportPath="Report\ConsSheet_OutP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
