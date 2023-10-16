<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="IAConsumablesOutStore.aspx.cs" Inherits="TyBus_Intranet_Test_V3.IAConsumablesOutStore" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="IAConsumablesOutStoreForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">電腦課耗材領用</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbSheetNo_Search" runat="server" CssClass="text-Right-Blue" Text="領用單號" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eSheetNoS_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eSheetNoE_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNoS_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNoS_Search_TextChanged" Width="40%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="～" Width="10%" />
                    <asp:TextBox ID="eDepNoE_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNoE_Search_TextChanged" Width="40%" />
                    <br />
                    <asp:Label ID="eDepNameS_Search" runat="server" CssClass="text-Left-Black" Width="50%" />
                    <asp:Label ID="eDepNameE_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbSheetStatus_Search" runat="server" CssClass="text-Right-Blue" Text="出貨情況" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlSheetStatus_Search" runat="server" CssClass="text-Left-Blus" Width="100%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuMan_Search" runat="server" CssClass="text-Right-Blue" Text="申領人" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eBuMan_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eBuMan_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eBuManName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuDate_Search" runat="server" CssClass="text-Right-Blue" Text="申領日期" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eBuDateS_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_3" runat="server" CssClass="text-Left-Blue" Text="～" />
                    <asp:TextBox ID="eBuDateE_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
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
                    <asp:DropDownList ID="ddlBrand_Search" runat="server" CssClass="text-Left-Black" Width="95%" AutoPostBack="True" DataSourceID="dsBrand_Search" DataTextField="ClassTxt" DataValueField="ClassNo" />
                    <asp:Label ID="eBrand_Search" runat="server" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCorrespondModel_Search" runat="server" CssClass="text-Right-Blue" Text="適用機型" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCorrespondModel_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
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
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:SqlDataSource ID="dsBrand_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        SelectCommand="SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材管理  fmIAConsumables Brand') ORDER BY ClassNo" />
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridIASheetA_List" runat="server" CssClass="text-Left-Blue" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="SheetNo" DataSourceID="dsIASheetA_List" ForeColor="#333333" GridLines="None" PageSize="5">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="SheetNo" HeaderText="出貨單號" ReadOnly="True" SortExpression="SheetNo" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="申請單位" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="BuDate" DataFormatString="{0:D}" HeaderText="申請日期" SortExpression="BuDate" />
                <asp:BoundField DataField="AssignManName" HeaderText="申請人" SortExpression="AssignManName" />
                <asp:BoundField DataField="SheetNote" HeaderText="摘要" SortExpression="SheetNote" />
                <asp:BoundField DataField="SheetStatus" HeaderText="SheetStatus" SortExpression="SheetStatus" Visible="False" />
                <asp:BoundField DataField="SheetStatus_C" HeaderText="處理情況" ReadOnly="True" SortExpression="SheetStatus_C" />
                <asp:BoundField DataField="RemarkA" HeaderText="備註" SortExpression="RemarkA" />
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
        <asp:FormView ID="fvIASheetA_Detail" runat="server" CssClass="text-Left-Black" DataKeyNames="SheetNo" DataSourceID="dsIASheetA_Detail"
            OnDataBound="fvIASheetA_Detail_DataBound" Width="100%">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CssClass="button-Blue" CausesValidation="True" OnClick="bbOK_Edit_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetNo_Edit" runat="server" CssClass="text-Right-Blue" Text="領用單號" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSheetNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="單位" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Width="35%" />
                            <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignMan_Edit" runat="server" CssClass="text-Right-Blue" Text="領料人" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="eAssignMan_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignMan_Edit_TextChanged" Text='<%# Eval("AssignMan") %>' Width="35%" />
                            <asp:Label ID="eAssignManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuMan_Edit" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                            <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_Edit" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                            <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="出貨狀況" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSheetStatus_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="2">
                            <asp:Label ID="lbSheetNote_Edit" runat="server" CssClass="text-Right-Blue" Text="摘要" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5" rowspan="2">
                            <asp:TextBox ID="eSheetNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("SheetNote") %>' Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStatusDate_Edit" runat="server" CssClass="text-Right-Blue" Text="狀態變動日" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStatusDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemarkA_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="100%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemarkA_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                    </tr>
                </table>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CssClass="button-Blue" CausesValidation="True" OnClick="bbOK_INS_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetNo_INS" runat="server" CssClass="text-Right-Blue" Text="領用單號" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSheetNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="單位" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Width="35%" />
                            <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignMan_INS" runat="server" CssClass="text-Right-Blue" Text="領料人" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="eAssignMan_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignMan_INS_TextChanged" Text='<%# Eval("AssignMan") %>' Width="35%" />
                            <asp:Label ID="eAssignManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuMan_INS" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                            <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_INS" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                            <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetStatus_INS" runat="server" CssClass="text-Right-Blue" Text="出貨狀況" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSheetStatus_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="2">
                            <asp:Label ID="lbSheetNote_INS" runat="server" CssClass="text-Right-Blue" Text="摘要" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5" rowspan="2">
                            <asp:TextBox ID="eSheetNote_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("SheetNote") %>' Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStatusDate_INS" runat="server" CssClass="text-Right-Blue" Text="狀態變動日" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStatusDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemarkA_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="100%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemarkA_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                    </tr>
                </table>
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNewA_Empty" runat="server" CssClass="button-Blue" CausesValidation="false" Text="新增" CommandName="New" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNewA_List" runat="server" CssClass="button-Black" CausesValidation="False" CommandName="New" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEditA_List" runat="server" CssClass="button-Blue" CausesValidation="False" CommandName="Edit" Text="修改" Width="120px" />
                &nbsp;<asp:Button ID="bbDeleteA_List" runat="server" CssClass="button-Red" CausesValidation="False" OnClick="bbDeleteA_List_Click" Text="刪除" Width="120px" />
                &nbsp;<asp:Button ID="bbOutStoreA_List" runat="server" CssClass="button-Blue" CausesValidation="false" OnClick="bbOutStoreA_List_Click" Text="出貨 (列印簽收單)" Width="120px" />
                &nbsp;<asp:Button ID="bbAbortA_List" runat="server" CssClass="button-Red" CausesValidation="false" OnClick="bbAbortA_List_Click" Text="本單作廢" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetNo_List" runat="server" CssClass="text-Right-Blue" Text="領用單號" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSheetNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="單位" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignMan_List" runat="server" CssClass="text-Right-Blue" Text="領料人" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eAssignMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' Width="35%" />
                            <asp:Label ID="eAssignManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuMan_List" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                            <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_List" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                            <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetStatus_List" runat="server" CssClass="text-Right-Blue" Text="出貨狀況" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSheetStatus_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                            <asp:Label ID="eSheetStatus_List" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="2">
                            <asp:Label ID="lbSheetNote_List" runat="server" CssClass="text-Right-Blue" Text="摘要" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5" rowspan="2">
                            <asp:TextBox ID="eSheetNote_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("SheetNote") %>' Width="97%" Height="97%" Enabled="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStatusDate_List" runat="server" CssClass="text-Right-Blue" Text="狀態變動日" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStatusDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemarkA_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="100%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemarkA_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Height="97%" Width="97%" Enabled="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
        <asp:SqlDataSource ID="dsIASheetA_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
            SelectCommand="SELECT SheetNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuManName, SheetNote, SheetStatus, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材進出單fmIASheetA      SheetStatus') AND (CLASSNO = a.SheetStatus)) AS SheetStatus_C, RemarkA, AssignMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.AssignMan)) AS AssignManName FROM IASheetA AS a WHERE (ISNULL(SheetNo, '') = '')" />
        <asp:SqlDataSource ID="dsIASheetA_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
            DeleteCommand="delete IASheetA where SheetNo = @SheetNo"
            InsertCommand="insert into IASheetA(SheetNo, SheetMode, BuDate, BuMan, SheetNote, SheetStatus, StatusDate, RemarkA, DepNo, AssignMan) values(@SheetNo, 'SO', GetDate(), @BuMan, @SheetNote, '000', GetDate(), @RemarkA, @DepNo, @AssignMan)"
            SelectCommand="SELECT SheetNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuManName, SheetNote, SheetStatus, StatusDate, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材進出單fmIASheetA      SheetStatus') AND (CLASSNO = a.SheetStatus)) AS SheetStatus_C, RemarkA, AssignMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = a.AssignMan)) AS AssignManNAme FROM IASheetA AS a WHERE (SheetNo = @SheetNo)"
            UpdateCommand="UPDATE IASheetA SET SheetNote = @SheetNote, RemarkA = @RemarkA, DepNo = @DepNo, ModifyDate = GETDATE(), ModifyMan = @ModifyMan, AssignMan = @AssignMan WHERE (SheetNo = @SheetNo)">
            <DeleteParameters>
                <asp:Parameter Name="SheetNo" />
            </DeleteParameters>
            <InsertParameters>
                <asp:Parameter Name="SheetNo" />
                <asp:Parameter Name="BuMan" />
                <asp:Parameter Name="SheetNote" />
                <asp:Parameter Name="RemarkA" />
                <asp:Parameter Name="DepNo" />
                <asp:Parameter Name="AssignMan" />
            </InsertParameters>
            <SelectParameters>
                <asp:ControlParameter ControlID="gridIASheetA_List" Name="SheetNo" PropertyName="SelectedValue" />
            </SelectParameters>
            <UpdateParameters>
                <asp:Parameter Name="SheetNote" />
                <asp:Parameter Name="RemarkA" />
                <asp:Parameter Name="DepNo" />
                <asp:Parameter Name="ModifyMan" />
                <asp:Parameter Name="AssignMan" />
                <asp:Parameter Name="SheetNo" />
            </UpdateParameters>
        </asp:SqlDataSource>
    </asp:Panel>
    <asp:Panel ID="plShowData_Detail" runat="server" CssClass="ShowPanel-Detail">
        <asp:GridView ID="gridIASheetB_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px" CellPadding="4" DataKeyNames="SheetNoItems" DataSourceID="dsIASheetB_List" PageSize="5" Width="100%">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="SheetNoItems" HeaderText="SheetNoItems" ReadOnly="True" SortExpression="SheetNoItems" Visible="False" />
                <asp:BoundField DataField="SheetNo" HeaderText="SheetNo" SortExpression="SheetNo" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                <asp:BoundField DataField="ConsNo" HeaderText="料號" SortExpression="ConsNo" />
                <asp:BoundField DataField="ConsName" HeaderText="品名" SortExpression="ConsName" />
                <asp:BoundField DataField="Unit" HeaderText="單位" SortExpression="Unit" Visible="False" />
                <asp:BoundField DataField="Unit_C" HeaderText="單位" ReadOnly="True" SortExpression="Unit_C" />
                <asp:BoundField DataField="Quantity" HeaderText="數量" SortExpression="Quantity" />
                <asp:BoundField DataField="RemarkB" HeaderText="項次摘要" SortExpression="RemarkB" />
                <asp:BoundField DataField="QtyMode" HeaderText="QtyMode" SortExpression="QtyMode" Visible="False" />
                <asp:BoundField DataField="ItemStatus" HeaderText="ItemStatus" SortExpression="ItemStatus" Visible="False" />
                <asp:BoundField DataField="ItemStatus_C" HeaderText="項次進度" ReadOnly="True" SortExpression="ItemStatus_C" />
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
        <asp:FormView ID="fvIASheetB_Detail" runat="server" DataKeyNames="SheetNoItems" DataSourceID="dsIASheetB_Detail" Width="100%" OnDataBound="fvIASheetB_Detail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOKB_Edit" runat="server" CssClass="button-Blue" CausesValidation="True" OnClick="bbOKB_Edit_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancelB_Edit" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItems_Edit" runat="server" CssClass="text-Right-Blue" Text="項次" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="100%" />
                            <asp:Label ID="eSheetNoItems_Edit" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                            <asp:Label ID="eSheetNoB_Edit" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsNo_Edit" runat="server" CssClass="text-Right-Blue" Text="料號" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5">
                            <asp:TextBox ID="eConsNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="15%" />
                            <asp:Button ID="bbGetConsNo_Edit" runat="server" CssClass="button-Black" Text="..." />
                            <asp:Label ID="eConsName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="70%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbQuantity_Edit" runat="server" CssClass="text-Right-Blue" Text="數量" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eQuantity_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Quantity") %>' Width="65%" />
                            <asp:Label ID="eUnit_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Unit_C") %>' Width="30%" />
                            <asp:Label ID="eUnit_Edit" runat="server" Text='<%# Eval("Unit") %>' Visible="false" />
                            <asp:Label ID="eQtyMode_Edit" runat="server" Text='<%# Eval("QtyMode") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="2">
                            <asp:Label ID="lbRemarkB_Edit" runat="server" CssClass="text-Right-Blue" Text="明細摘要" Width="100%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-8Col" rowspan="2" colspan="5">
                            <asp:TextBox ID="eRemarkB_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItemStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="明細進度" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItemStatus_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="100%" />
                            <asp:Label ID="eItemStatus_Edit" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuMan_Edit" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                            <asp:Label ID="eBuMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_Edit" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                            <asp:Label ID="eModifyMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="100%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                    </tr>
                </table>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOKB_INS" runat="server" CssClass="button-Blue" CausesValidation="True" OnClick="bbOKB_INS_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancelB_INS" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItems_INS" runat="server" CssClass="text-Right-Blue" Text="項次" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItems_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="100%" />
                            <asp:Label ID="eSheetNoItems_INS" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                            <asp:Label ID="eSheetNoB_INS" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsNo_INS" runat="server" CssClass="text-Right-Blue" Text="料號" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5">
                            <asp:TextBox ID="eConsNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="15%" />
                            <asp:Button ID="bbGetConsNo_INS" runat="server" CssClass="button-Black" Text="..." />
                            <asp:Label ID="eConsName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="70%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbQuantity_INS" runat="server" CssClass="text-Right-Blue" Text="數量" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eQuantity_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Quantity") %>' Width="65%" />
                            <asp:Label ID="eUnit_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Unit_C") %>' Width="30%" />
                            <asp:Label ID="eUnit_INS" runat="server" Text='<%# Eval("Unit") %>' Visible="false" />
                            <asp:Label ID="eQtyMode_INS" runat="server" Text='<%# Eval("QtyMode") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="2">
                            <asp:Label ID="lbRemarkB_INS" runat="server" CssClass="text-Right-Blue" Text="明細摘要" Width="100%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-8Col" rowspan="2" colspan="5">
                            <asp:TextBox ID="eRemarkB_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItemStatus_INS" runat="server" CssClass="text-Right-Blue" Text="明細進度" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItemStatus_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="100%" />
                            <asp:Label ID="eItemStatus_INS" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuMan_INS" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                            <asp:Label ID="eBuMan_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_INS" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                            <asp:Label ID="eModifyMan_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="100%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                    </tr>
                </table>
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNewB_Empty" runat="server" CssClass="text-Left-Blue" CommandName="New" Text="新增明細" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNewB_List" runat="server" CssClass="button-Black" CausesValidation="False" CommandName="New" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEditB_List" runat="server" CssClass="button-Blue" CausesValidation="False" CommandName="Edit" Text="修改" Width="120px" />
                &nbsp;<asp:Button ID="bbDeleteB_List" runat="server" CssClass="button-Red" CausesValidation="False" OnClick="bbDeleteB_List_Click" Text="刪除" Width="120px" />
                &nbsp;<asp:Button ID="bbOutStoreB_List" runat="server" CssClass="button-Black" CausesValidation="false" OnClick="bbOutStoreB_List_Click" Text="明細出貨" Width="120px" />
                &nbsp;<asp:Button ID="bbArrivedB_List" runat="server" CssClass="button-Blue" CausesValidation="false" OnClick="bbArrivedB_List_Click" Text="明細簽收" Width="120px" />
                &nbsp;<asp:Button ID="bbAbortB_List" runat="server" CssClass="button-Red" CausesValidation="false" OnClick="bbAbortB_List_Click" Text="明細作廢" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItems_List" runat="server" CssClass="text-Right-Blue" Text="項次" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="100%" />
                            <asp:Label ID="eSheetNoItems_List" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                            <asp:Label ID="eSheetNoB_List" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbConsNo_List" runat="server" CssClass="text-Right-Blue" Text="料號" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5">
                            <asp:Label ID="eConsNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="15%" />
                            <asp:Label ID="eConsName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="80%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbQuantity_List" runat="server" CssClass="text-Right-Blue" Text="數量" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eQuantity_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Quantity") %>' Width="65%" />
                            <asp:Label ID="eUnit_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Unit_C") %>' Width="30%" />
                            <asp:Label ID="eUnit_List" runat="server" Text='<%# Eval("Unit") %>' Visible="false" />
                            <asp:Label ID="eQtyMode_List" runat="server" Text='<%# Eval("QtyMode") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="2">
                            <asp:Label ID="lbRemarkB_List" runat="server" CssClass="text-Right-Blue" Text="明細摘要" Width="100%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-8Col" rowspan="2" colspan="5">
                            <asp:TextBox ID="eRemarkB_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItemStatus_List" runat="server" CssClass="text-Right-Blue" Text="明細進度" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItemStatus_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="100%" />
                            <asp:Label ID="eItemStatus_List" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuMan_List" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                            <asp:Label ID="eBuMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_List" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                            <asp:Label ID="eModifyMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="100%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColWidth-8Col" />
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
        <asp:SqlDataSource ID="dsIASheetB_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
            SelectCommand="SELECT b.SheetNoItems, b.SheetNo, b.Items, b.ConsNo, c.ConsName, c.Unit, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材管理  fmIAConsumables Unit') AND (CLASSNO = c.Unit)) AS Unit_C, b.Quantity, b.RemarkB, b.QtyMode, b.ItemStatus, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (FKEY = '電腦課耗材進出單fmIASheetB      ItemStatus') AND (CLASSNO = b.ItemStatus)) AS ItemStatus_C FROM IASheetB AS b LEFT OUTER JOIN IAConsumables AS c ON c.ConsNo = b.ConsNo WHERE (b.SheetNo = @SheetNo)">
            <SelectParameters>
                <asp:ControlParameter ControlID="gridIASheetA_List" Name="SheetNo" PropertyName="SelectedValue" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="dsIASheetB_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
            DeleteCommand="delete IASheetB where SheetNoItems = @SheetNoItems"
            InsertCommand="INSERT INTO IASheetB(SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, QtyMode, ItemStatus, BuMan, BuDate) VALUES (@SheetNoItems, @SheetNo, @Items, @ConsNo, @Price, @Quantity, @Amount, @RemarkB, @QtyMode, @ItemStatus, @BuMan, @BuDate)"
            SelectCommand="SELECT b.SheetNoItems, b.SheetNo, b.Items, b.ConsNo, c.ConsName, c.Unit, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材管理  fmIAConsumables Unit') AND (CLASSNO = c.Unit)) AS Unit_C, b.Quantity, b.RemarkB, b.QtyMode, b.ItemStatus, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (FKEY = '電腦課耗材進出單fmIASheetB      ItemStatus') AND (CLASSNO = b.ItemStatus)) AS ItemStatus_C, b.BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = b.BuMan)) AS BuMan_C, b.BuDate, b.ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = b.ModifyMan)) AS ModifyMan_C, b.ModifyDate FROM IASheetB AS b LEFT OUTER JOIN IAConsumables AS c ON c.ConsNo = b.ConsNo WHERE (b.SheetNoItems = @SheetNoItems)"
            UpdateCommand="UPDATE IASheetB SET ConsNo = @ConsNo, Price = @Price, Quantity = @Quantity, Amount = @Amount, RemarkB = @RemarkB, QtyMode = @QtyMode, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate WHERE (SheetNoItems = @SheetNoItems)">
            <DeleteParameters>
                <asp:Parameter Name="SheetNoItems" />
            </DeleteParameters>
            <InsertParameters>
                <asp:Parameter Name="SheetNoItems" />
                <asp:Parameter Name="SheetNo" />
                <asp:Parameter Name="Items" />
                <asp:Parameter Name="ConsNo" />
                <asp:Parameter Name="Price" />
                <asp:Parameter Name="Quantity" />
                <asp:Parameter Name="Amount" />
                <asp:Parameter Name="RemarkB" />
                <asp:Parameter Name="QtyMode" />
                <asp:Parameter Name="ItemStatus" />
                <asp:Parameter Name="BuMan" />
                <asp:Parameter Name="BuDate" />
            </InsertParameters>
            <SelectParameters>
                <asp:ControlParameter ControlID="gridIASheetB_List" Name="SheetNoItems" PropertyName="SelectedValue" />
            </SelectParameters>
            <UpdateParameters>
                <asp:Parameter Name="ConsNo" />
                <asp:Parameter Name="Price" />
                <asp:Parameter Name="Quantity" />
                <asp:Parameter Name="Amount" />
                <asp:Parameter Name="RemarkB" />
                <asp:Parameter Name="QtyMode" />
                <asp:Parameter Name="ModifyMan" />
                <asp:Parameter Name="ModifyDate" />
                <asp:Parameter Name="SheetNoItems" />
            </UpdateParameters>
        </asp:SqlDataSource>
    </asp:Panel>
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
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
            <LocalReport ReportPath="Report\IAConsumablesOutStoreP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
