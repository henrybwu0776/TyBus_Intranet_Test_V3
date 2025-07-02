<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ConsSheet_Pur.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ConsSheet_Pur" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="ConsSheet_PurForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">總務耗材採購單</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <div>
            <asp:Label ID="eErrorMSG_Main" runat="server" CssClass="errorMessageText" Visible="false" />
        </div>
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuDate_Search" runat="server" CssClass="text-Right-Blue" Text="申請日期" Width="95%" />
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
        <asp:GridView ID="gridConsSheetA_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="SheetNo" DataSourceID="sdsConsSheetA_List" ForeColor="#333333" GridLines="None" PageSize="5" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" ItemStyle-Width="60px" ItemStyle-CssClass="fixedWidth" />
                <asp:BoundField DataField="SheetNo" HeaderText="單號" ReadOnly="True" SortExpression="SheetNo" ItemStyle-Width="80px" ItemStyle-CssClass="fixedWidth" />
                <asp:BoundField DataField="BuDate" HeaderText="建檔日期" SortExpression="BuDate" DataFormatString="{0:d}" />
                <asp:BoundField DataField="BuMan_C" HeaderText="建檔人" SortExpression="BuMan_C" />
                <asp:BoundField DataField="SupName" HeaderText="供應廠商" SortExpression="SupName" />
                <asp:BoundField DataField="TotalAmount" HeaderText="總金額" SortExpression="TotalAmount" />
                <asp:BoundField DataField="SheetStatus_C" HeaderText="處理進度" SortExpression="SheetStatus_C" />
            </Columns>
            <EditRowStyle BackColor="#2461BF" />
            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#EFF3FB" />
            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#F5F7FB" />
            <SortedAscendingHeaderStyle BackColor="#6D95E1" />
            <SortedDescendingCellStyle BackColor="#E9EBEF" />
            <SortedDescendingHeaderStyle BackColor="#4870BE" />
        </asp:GridView>
        <asp:FormView ID="fvConsSheetA_Detail" runat="server" DataKeyNames="SheetNo" DataSourceID="sdsConsSheetA_Detail" OnDataBound="fvConsSheetA_Detail_DataBound" Width="100%">
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
                            <asp:Label ID="lbAmountA_Edit" runat="server" CssClass="text-Right-Blue" Text="金額小計" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAmountA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTaxType_Edit" runat="server" CssClass="text-Right-Blue" Text="稅別" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:DropDownList ID="ddlTaxType_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlTaxType_Edit_SelectedIndexChanged" Width="95%" />
                            <asp:Label ID="eTaxType_Edit" runat="server" Text='<%# Eval("TaxType") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTaxRate_Edit" runat="server" CssClass="text-Right-Blue" Text="稅率" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eTaxRate_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eTaxRate_Edit_TextChanged" Text='<%# Eval("TaxRate") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTaxAMT_Edit" runat="server" CssClass="text-Right-Blue" Text="稅額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eTaxAMT_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TaxAMT") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTotalAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="總金額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eTotalAmount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TotalAmount") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPayMode_Edit" runat="server" CssClass="text-Right-Blue" Text="付款方式" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:DropDownList ID="ddlPayMode_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlPayMode_Edit_SelectedIndexChanged" Width="95%" />
                            <asp:Label ID="ePayMode_Edit" runat="server" Text='<%# Eval("PayMode") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPayDate_Edit" runat="server" CssClass="text-Right-Blue" Text="付款日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePayDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PayDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
                            <asp:Label ID="lbAmountA_INS" runat="server" CssClass="text-Right-Blue" Text="金額小計" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAmountA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTaxType_INS" runat="server" CssClass="text-Right-Blue" Text="稅別" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:DropDownList ID="ddlTaxType_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlTaxType_INS_SelectedIndexChanged" Width="95%" />
                            <asp:Label ID="eTaxType_INS" runat="server" Text='<%# Eval("TaxType") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTaxRate_INS" runat="server" CssClass="text-Right-Blue" Text="稅率" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eTaxRate_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eTaxRate_INS_TextChanged" Text='<%# Eval("TaxRate") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTaxAMT_INS" runat="server" CssClass="text-Right-Blue" Text="稅額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eTaxAMT_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TaxAMT") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTotalAmount_INS" runat="server" CssClass="text-Right-Blue" Text="總金額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eTotalAmount_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TotalAmount") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPayMode_INS" runat="server" CssClass="text-Right-Blue" Text="付款方式" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:DropDownList ID="ddlPayMode_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlPayMode_INS_SelectedIndexChanged" Width="95%" />
                            <asp:Label ID="ePayMode_INS" runat="server" Text='<%# Eval("PayMode") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPayDate_INS" runat="server" CssClass="text-Right-Blue" Text="付款日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePayDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PayDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
                <asp:Button ID="bbNewA_List" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
                <asp:Button ID="bbEditA_List" runat="server" CausesValidation="false" CssClass="button-Blue" CommandName="Edit" Text="修改" Width="120px" />
                <asp:Button ID="bbArriveA_List" runat="server" CausesValidation="false" CssClass="button-Black" OnClick="bbArriveA_List_Click" Text="全單到貨" Width="120px" />
                <asp:Button ID="bbCloseA_List" runat="server" CausesValidation="false" CssClass="button-Blue" OnClick="bbCloseA_List_Click" Text="付款結案" Width="120px" />
                <asp:Button ID="bbPtint_List" runat="server" CausesValidation="false" CssClass="button-Black" OnClick="bbPtint_List_Click" Text="列印採購單" Width="120px" />
                <asp:Button ID="bbAbortA_List" runat="server" CausesValidation="false" CssClass="button-Blue" OnClick="bbAbortA_List_Click" Text="採購單作廢" Width="120px" />
                <asp:Button ID="bbDelA_List" runat="server" CausesValidation="false" CssClass="button-Red" OnClick="bbDelA_List_Click" Text="刪除採購單" Width="120px" />
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
                            <asp:Label ID="lbAmountA_List" runat="server" CssClass="text-Right-Blue" Text="金額小計" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAmountA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTaxType_List" runat="server" CssClass="text-Right-Blue" Text="稅別" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eTaxType_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TaxType_C") %>' Width="95%" />
                            <asp:Label ID="eTaxType_List" runat="server" Text='<%# Eval("TaxType") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTaxRate_List" runat="server" CssClass="text-Right-Blue" Text="稅率" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eTaxRate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TaxRate") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTaxAMT_List" runat="server" CssClass="text-Right-Blue" Text="稅額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eTaxAMT_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TaxAMT") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTotalAmount_List" runat="server" CssClass="text-Right-Blue" Text="總金額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eTotalAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TotalAmount") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPayMode_List" runat="server" CssClass="text-Right-Blue" Text="付款方式" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePayMode_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PayMode_C") %>' Width="95%" />
                            <asp:Label ID="ePayMode_List" runat="server" Text='<%# Eval("PayMode") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPayDate_List" runat="server" CssClass="text-Right-Blue" Text="付款日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePayDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PayDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
        <asp:SqlDataSource ID="sdsConsSheetA_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select a.SheetNo, a.SheetMode, d.ClassTxt SheetMode_C, a.SupNo, c.[Name] SupName, a.BuDate, a.BuMan, e2.[Name] BuMan_C, a.Amount, 
       a.TaxRate, a.TaxType, d2.ClassTxt TaxType_C, a.TaxAMT, a.TotalAmount, a.SheetStatus, d3.ClassTxt SheetStatus_C, a.StatusDate, 
	   a.PayDate, a.PayMode, d4.ClassTxt PayMode_C, a.RemarkA, a.ModifyMan, e3.[Name] ModifyMan_C, a.ModifyDate, 
	   a.SheetNote 
  from ConsSheetA a left  join DBDICB d on d.ClassNo = a.SheetMode and d.FKey = '總務課耗材進出單ConsSheetA      SheetMode'
                    left join DBDICB d2 on d2.ClassNo = a.TaxType and d2.FKey = '總務請購單      ConsSheetA      TAXTYPE'
                    left join DBDICB d3 on d3.ClassNo = a.SheetStatus and d3.FKey = '總務耗材進出單  fmConsSheetA    SheetStatus'
                    left join DBDICB d4 on d4.ClassNo = a.PayMode and d4.FKey = '總務請購單      ConsSheetA      PayMode'
                    left join Employee e2 on e2.EmpNo = a.BuMan
                    left join Employee e3 on e3.EmpNo = a.ModifyMan
                    left join [Custom] c on c.Code = a.SupNo and c.[Types] = 'S'
 where SheetNo = @SheetNo">
            <SelectParameters>
                <asp:ControlParameter ControlID="gridConsSheetA_List" Name="SheetNo" PropertyName="SelectedValue" />
            </SelectParameters>
        </asp:SqlDataSource>
    </asp:Panel>
    <asp:Panel ID="plShowDataB" runat="server" CssClass="ShowPanel-Detail">
        <asp:GridView ID="gridConsSheetB_List" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" BorderStyle="None" CellPadding="4" DataKeyNames="SheetNoItems" DataSourceID="sdsConsSheetB_List" ForeColor="#333333" GridLines="None" PageSize="5">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" ItemStyle-Width="60px" ItemStyle-CssClass="fixedWidth" />
                <asp:BoundField DataField="SheetNoItems" HeaderText="SheetNoItems" ReadOnly="True" SortExpression="SheetNoItems" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" ItemStyle-Width="70px" ItemStyle-CssClass="fixedWidth">
                    <ItemStyle CssClass="fixedWidth" Width="70px" />
                </asp:BoundField>
                <asp:BoundField DataField="ConsName" HeaderText="品名" SortExpression="ConsName" />
                <asp:BoundField DataField="Price" HeaderText="單價" SortExpression="Price" />
                <asp:BoundField DataField="Quantity" HeaderText="數量" SortExpression="Quantity" />
                <asp:BoundField DataField="ConsUnit_C" HeaderText="單位" SortExpression="ConsUnit_C" />
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
        <asp:FormView ID="fvConsSheetB_Detail" runat="server" Width="100%" DataKeyNames="SheetNoItems" DataSourceID="sdsConsSheetB_Detail" OnDataBound="fvConsSheetB_Detail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOKB_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKB_Edit_Click" Text="確定" Width="120px" />
                <asp:Button ID="bbCancelB_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItems_Edit" runat="server" CssClass="text-Right-Blue" Text="項次" Width="95%" />
                            <asp:Label ID="eSheetNoItems_Edit" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoB_Edit" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbConsNo_Edit" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="5">
                            <asp:TextBox ID="eConsNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="25%" />
                            <asp:TextBox ID="eConsName_Edit" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("ConsName") %>' Width="65%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItemStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="處理進度" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItemStatus_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                            <asp:Label ID="eItemStatus_Edit" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPrice_Edit" runat="server" CssClass="text-Right-Blue" Text="單價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePrice_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="ePrice_Edit_TextChanged" Text='<%# Eval("Price") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3">
                            <asp:Label ID="lbRemarkB_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3" colspan="7">
                            <asp:TextBox ID="eRemarkB_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbQuantity_Edit" runat="server" CssClass="text-Right-Blue" Text="數量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eQuantity_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="ePrice_Edit_TextChanged" Text='<%# Eval("Quantity") %>' Width="40%" />
                            <asp:DropDownList ID="ddlConsUnit_Edit" runat="server" CssClass="text-Left-Black" OnSelectedIndexChanged="ddlConsUnit_Edit_SelectedIndexChanged" Width="40%" />
                            <asp:TextBox ID="eConsUnit_Edit" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAmountB_Edit" runat="server" CssClass="text-Right-Blue" Text="小計" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAmountB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuManB_Edit" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuMan_CB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateB_Edit" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuDateB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManB_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyMan_CB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateB_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
                <asp:Button ID="bbOKB_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKB_INS_Click" Text="確定" Width="120px" />
                <asp:Button ID="bbCancelB_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItems_INS" runat="server" CssClass="text-Right-Blue" Text="項次" Width="95%" />
                            <asp:Label ID="eSheetNoItems_INS" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItems_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoB_INS" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbConsNo_INS" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="5">
                            <asp:TextBox ID="eConsNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="25%" />
                            <asp:TextBox ID="eConsName_INS" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("ConsName") %>' Width="65%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItemStatus_INS" runat="server" CssClass="text-Right-Blue" Text="處理進度" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItemStatus_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                            <asp:Label ID="eItemStatus_INS" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPrice_INS" runat="server" CssClass="text-Right-Blue" Text="單價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePrice_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="ePrice_INS_TextChanged" Text='<%# Eval("Price") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3">
                            <asp:Label ID="lbRemarkB_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3" colspan="7">
                            <asp:TextBox ID="eRemarkB_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbQuantity_INS" runat="server" CssClass="text-Right-Blue" Text="數量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eQuantity_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="ePrice_INS_TextChanged" Text='<%# Eval("Quantity") %>' Width="40%" />
                            <asp:DropDownList ID="ddlConsUnit_INS" runat="server" CssClass="text-Left-Black" OnSelectedIndexChanged="ddlConsUnit_INS_SelectedIndexChanged" Width="40%" />
                            <asp:TextBox ID="eConsUnit_INS" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAmountB_INS" runat="server" CssClass="text-Right-Blue" Text="小計" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAmountB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuManB_INS" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuMan_CB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateB_INS" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuDateB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManB_INS" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyMan_CB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateB_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
                <asp:Button ID="bbNewB_Empty" runat="server" CssClass="button-Black" CausesValidation="true" CommandName="New" Text="單筆新增明細" Width="120px" />
                <asp:Button ID="bbExportDetail_Empty" runat="server" CssClass="button-Blue" CausesValidation="false" OnClick="bbExportDetail_List_Click" Text="請購匯入明細" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNewB_List" runat="server" CssClass="button-Black" CausesValidation="true" CommandName="New" Text="單筆新增明細" Width="120px" />
                <asp:Button ID="bbExportDetail_List" runat="server" CssClass="button-Blue" CausesValidation="false" OnClick="bbExportDetail_List_Click" Text="請購匯入明細" Width="120px" />
                <asp:Button ID="bbEditB_List" runat="server" CssClass="button-Black" CausesValidation="false" CommandName="Edit" Text="修改明細" Width="120px" />
                <asp:Button ID="bbAriveB_List" runat="server" CssClass="button-Blue" CausesValidation="false" OnClick="bbAriveB_List_Click" Text="明細到貨" Width="120px" />
                <asp:Button ID="bbAbortB_List" runat="server" CssClass="button-Black" CausesValidation="false" OnClick="bbAbortB_List_Click" Text="明細作廢" Width="120px" />
                <asp:Button ID="bbDeleteB_List" runat="server" CssClass="button-Red" CausesValidation="false" OnClick="bbDeleteB_List_Click" Text="刪除明細" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItems_List" runat="server" CssClass="text-Right-Blue" Text="項次" Width="95%" />
                            <asp:Label ID="eSheetNoItems_List" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoB_List" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbConsNo_List" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="5">
                            <asp:Label ID="eConsNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="25%" />
                            <asp:Label ID="eConsName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="65%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItemStatus_List" runat="server" CssClass="text-Right-Blue" Text="處理進度" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItemStatus_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                            <asp:Label ID="eItemStatus_List" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPrice_List" runat="server" CssClass="text-Right-Blue" Text="單價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePrice_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Price") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3">
                            <asp:Label ID="lbRemarkB_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3" colspan="7">
                            <asp:TextBox ID="eRemarkB_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Enabled="false" Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbQuantity_List" runat="server" CssClass="text-Right-Blue" Text="數量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eQuantity_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Quantity") %>' Width="40%" />
                            <asp:Label ID="eConsUnit_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="40%" />
                            <asp:Label ID="eConsUnit_List" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAmountB_List" runat="server" CssClass="text-Right-Blue" Text="小計" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAmountB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuManB_List" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuMan_CB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateB_List" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuDateB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManB_List" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyMan_CB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateB_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
        <asp:SqlDataSource ID="sdsConsSheetB_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select b.SheetNoItems, b.Items, c.ConsName, b.Price, b.Quantity, d.ClassTxt ConsUnit_C from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo left join DBDICB d on d.ClassNo = b.ConsUnit and d.FKey = '耗材庫存        CONSUMABLES     ConsUnit' where b.SheetNo = @SheetNo ">
            <SelectParameters>
                <asp:ControlParameter ControlID="gridConsSheetA_List" Name="SheetNo" PropertyName="SelectedValue" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="sdsConsSheetB_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select b.SheetNoItems, b.SheetNo, b.Items, b.ConsNo, c.ConsName, b.Price, b.Quantity, b.Amount, b.RemarkB, b.ItemStatus, d2.ClassTxt ItemStatus_C, 
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
    </asp:Panel>
    <asp:Panel ID="plPickupDetail" runat="server" CssClass="ShowPanel-Detail">
        <div>
            <asp:Button ID="bbSelectAll_Order" runat="server" CssClass="button-Blue" Text="全選" OnClick="bbSelectAll_Order_Click" Width="120px" />
            <asp:Button ID="bbUnselAll_Order" runat="server" CssClass="button-Black" Text="取消全選" OnClick="bbUnselAll_Order_Click" Width="120px" />
            <asp:Button ID="bbOKC_Order" runat="server" CssClass="button-Blue" Text="確定" OnClick="bbOKC_Order_Click" Width="120px" />
            <asp:Button ID="bbCancelC_Order" runat="server" CssClass="button-Red" Text="取消" OnClick="bbCancelC_Order_Click" Width="120px" />
        </div>
        <asp:GridView ID="gridPickup_List" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="SheetNoItems" DataSourceID="sdsPickup_List" GridLines="None" Width="100%">
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
                <asp:BoundField DataField="StockQty" HeaderText="現有庫存量" SortExpression="StockQty" />
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
        <asp:SqlDataSource ID="sdsPickup_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select b.SheetNoItems, b.SheetNo, b.Items, b.ConsNo, c.ConsName, b.Quantity, c.StockQty, b.RemarkB, b.Price
  from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo  where isnull(SheetNoItems, '') = ''"></asp:SqlDataSource>
    </asp:Panel>
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
        <div>
            <asp:Button ID="bbCloseReport" runat="server" CausesValidation="true" CssClass="button-Red" Text="關閉預覽" OnClick="bbCloseReport_Click" Width="120px" />
        </div>
        <rsweb:ReportViewer ID="rvPrint" runat="server" Width="100%" BackColor="" ClientIDMode="AutoID" HighlightBackgroundColor="" InternalBorderColor="204, 204, 204" InternalBorderStyle="Solid" InternalBorderWidth="1px" LinkActiveColor="" LinkActiveHoverColor="" LinkDisabledColor="" PrimaryButtonBackgroundColor="" PrimaryButtonForegroundColor="" PrimaryButtonHoverBackgroundColor="" PrimaryButtonHoverForegroundColor="" SecondaryButtonBackgroundColor="" SecondaryButtonForegroundColor="" SecondaryButtonHoverBackgroundColor="" SecondaryButtonHoverForegroundColor="" SplitterBackColor="" ToolbarDividerColor="" ToolbarForegroundColor="" ToolbarForegroundDisabledColor="" ToolbarHoverBackgroundColor="" ToolbarHoverForegroundColor="" ToolBarItemBorderColor="" ToolBarItemBorderStyle="Solid" ToolBarItemBorderWidth="1px" ToolBarItemHoverBackColor="" ToolBarItemPressedBorderColor="51, 102, 153" ToolBarItemPressedBorderStyle="Solid" ToolBarItemPressedBorderWidth="1px" ToolBarItemPressedHoverBackColor="153, 187, 226">
            <LocalReport ReportPath="Report\ConsSheet_PurP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
