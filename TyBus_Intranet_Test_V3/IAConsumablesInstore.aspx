<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="IAConsumablesInstore.aspx.cs" Inherits="TyBus_Intranet_Test_V3.IAConsumablesInstore" %>

<asp:Content ID="IAConsumablesInstoreForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">電腦課耗材進貨</a>
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
                    <asp:DropDownList ID="ddlSheetStatus_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="True" OnSelectedIndexChanged="ddlSheetStatus_Search_SelectedIndexChanged"
                        DataSourceID="dsSheetStatus_Search" DataTextField="ClassTxt" DataValueField="ClassNo" Width="95%" />
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
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridIASheetList" runat="server" BackColor="White" BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" CellPadding="3" GridLines="Horizontal" Width="100%"
            AllowPaging="True" AutoGenerateColumns="False" DataKeyNames="SheetNo" DataSourceID="dsIASheetA_List" PageSize="5">
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
        <asp:FormView ID="fvIASheetA_Detail" runat="server" Width="100%" DataKeyNames="SheetNo" DataSourceID="dsIASheetA_Detail" OnDataBound="fvIASheetA_Detail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_Edit_Click" Text="確定" Width="120px" />
                &nbsp;
                <asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbSheetNo_Edit" runat="server" CssClass="text-Right-Blue" Text="進貨單號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eSheetNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbSupNo_Edit" runat="server" CssClass="text-Right-Blue" Text="供應商" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                                    <asp:TextBox ID="eSupNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupNo") %>' AutoPostBack="true" OnTextChanged="eSupNo_Edit_TextChanged" Width="35%" />
                                    <asp:TextBox ID="eSupName_Edit" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("SupName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbSheetStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="單據狀態" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eSheetStatus_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbStatusDate_Edit" runat="server" CssClass="text-Right-Blue" Text="到貨日" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eStatusDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="開單日" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="開單人員" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                                    <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbPayMode_Edit" runat="server" CssClass="text-Right-Blue" Text="付款方式" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:DropDownList ID="ddlPayMode_Edit" runat="server" CssClass="text-Left-Black" Width="95%" AutoPostBack="True"
                                        DataSourceID="dsPayMode_Edit" DataTextField="ClassTxt" DataValueField="ClassNo" OnSelectedIndexChanged="ddlPayMode_Edit_SelectedIndexChanged" />
                                    <asp:Label ID="ePayMode_Edit" runat="server" Text='<%# Eval("PayMode") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbPayDate_Edit" runat="server" CssClass="text-Right-Blue" Text="付款日" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="ePayDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PayDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_High ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbSheetNote_Edit" runat="server" CssClass="text-Right-Blue" Text="摘要" Width="100%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="9">
                                    <asp:TextBox ID="eSheetNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("SheetNote") %>' Width="95%" Height="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="小計金額" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eAmount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbTaxType_Edit" runat="server" CssClass="text-Right-Blue" Text="稅別" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:DropDownList ID="ddlTaxType_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="True" OnSelectedIndexChanged="ddlTaxType_Edit_SelectedIndexChanged" Width="95%"
                                        DataSourceID="dsTaxType_Edit" DataTextField="ClassTxt" DataValueField="ClassNo">
                                    </asp:DropDownList>
                                    <asp:Label ID="eTaxType_Edit" runat="server" Text='<%# Eval("TaxType") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbTaxRate_Edit" runat="server" CssClass="text-Right-Blue" Text="稅率" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:TextBox ID="eTaxRate_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnTextChanged="ddlTaxType_Edit_SelectedIndexChanged" Text='<%# Eval("TaxRate") %>' Width="75%" />
                                    <asp:Label ID="lbSplit_3" runat="server" CssClass="text-Left-Black" Text="％" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbTaxAMT_Edit" runat="server" CssClass="text-Right-Blue" Text="稅額" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:TextBox ID="eTaxAMT_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnTextChanged="eTaxAMT_Edit_TextChanged" Text='<%# Eval("TaxAMT") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbTotalAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="總金額" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eTotalAmount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TotalAmount") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_High ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbRemarkA_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="100%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="9">
                                    <asp:TextBox ID="eRemarkA_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Width="95%" Height="97%" />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_INS_Click" Text="確定" Width="120px" />
                &nbsp;
                <asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbSheetNo_INS" runat="server" CssClass="text-Right-Blue" Text="進貨單號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eSheetNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbSupNo_INS" runat="server" CssClass="text-Right-Blue" Text="供應商" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                                    <asp:TextBox ID="eSupNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupNo") %>' Width="35%" />
                                    <asp:TextBox ID="eSupName_INS" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("SupName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbSheetStatus_INS" runat="server" CssClass="text-Right-Blue" Text="單據狀態" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eSheetStatus_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbStatusDate_INS" runat="server" CssClass="text-Right-Blue" Text="到貨日" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eStatusDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="開單日" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="開單人員" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                                    <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbPayMode_INS" runat="server" CssClass="text-Right-Blue" Text="付款方式" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:DropDownList ID="ddlPayMode_INS" runat="server" CssClass="text-Left-Black" Width="95%" AutoPostBack="True"
                                        DataSourceID="dsPayMode_INS" DataTextField="ClassTxt" DataValueField="ClassNo" OnSelectedIndexChanged="ddlPayMode_INS_SelectedIndexChanged" />
                                    <asp:Label ID="ePayMode_INS" runat="server" Text='<%# Eval("PayMode") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbPayDate_INS" runat="server" CssClass="text-Right-Blue" Text="付款日" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="ePayDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PayDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_High ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbSheetNote_INS" runat="server" CssClass="text-Right-Blue" Text="摘要" Width="100%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="9">
                                    <asp:TextBox ID="eSheetNote_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("SheetNote") %>' Width="95%" Height="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbAmount_INS" runat="server" CssClass="text-Right-Blue" Text="小計金額" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eAmount_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbTaxType_INS" runat="server" CssClass="text-Right-Blue" Text="稅別" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:DropDownList ID="ddlTaxType_INS" runat="server" CssClass="text-Left-Black"
                                        AutoPostBack="True" OnSelectedIndexChanged="ddlTaxType_INS_SelectedIndexChanged" Width="95%"
                                        DataSourceID="dsTaxType_INS" DataTextField="ClassTxt" DataValueField="ClassNo">
                                    </asp:DropDownList>
                                    <asp:Label ID="eTaxType_INS" runat="server" Text='<%# Eval("TaxType") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbTaxRate_INS" runat="server" CssClass="text-Right-Blue" Text="稅率" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:TextBox ID="eTaxRate_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnTextChanged="ddlTaxType_INS_SelectedIndexChanged" Text='<%# Eval("TaxRate") %>' Width="75%" />
                                    <asp:Label ID="lbSplit_4" runat="server" CssClass="text-Left-Black" Text="％" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbTaxAMT_INS" runat="server" CssClass="text-Right-Blue" Text="稅額" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:TextBox ID="eTaxAMT_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnTextChanged="eTaxAMT_INS_TextChanged" Text='<%# Eval("TaxAMT") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbTotalAmount_INS" runat="server" CssClass="text-Right-Blue" Text="總金額" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eTotalAmount_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TotalAmount") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_High ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbRemarkA_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="100%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="9">
                                    <asp:TextBox ID="eRemarkA_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Width="95%" Height="97%" />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CssClass="button-Black" Text="新增" CommandName="New" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_List" runat="server" CssClass="button-Black" Text="新增" CommandName="New" Width="120px" />
                &nbsp;
                <asp:Button ID="bbModify_List" runat="server" CssClass="button-Blue" Text="修改" CommandName="Edit" Width="120px" />
                &nbsp;
                <asp:Button ID="bbDelete_List" runat="server" CssClass="button-Red" Text="刪除" OnClick="bbDelete_List_Click" Width="120px" />
                &nbsp;
                <asp:Button ID="bbIsArrived_List" runat="server" CssClass="button-Black" Text="收貨" OnClick="bbIsArrived_List_Click" Width="120px" />
                &nbsp;
                <asp:Button ID="bbIsPaid_List" runat="server" CssClass="button-Blue" Text="付款" OnClick="bbIsPaid_List_Click" Width="120px" />
                &nbsp;
                <asp:Button ID="bbIsAbort_List" runat="server" CssClass="button-Red" Text="本單作廢" OnClick="bbIsAbort_List_Click" Width="120px" />
                &nbsp;
                <asp:Button ID="bbIsClosed_List" runat="server" CssClass="button-Blue" Text="結案" OnClick="bbIsClosed_List_Click" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbSheetNo_List" runat="server" CssClass="text-Right-Blue" Text="進貨單號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eSheetNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbSupNo_List" runat="server" CssClass="text-Right-Blue" Text="供應商" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                                    <asp:Label ID="eSupNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupNo") %>' Width="35%" />
                                    <asp:Label ID="eSupName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbSheetStatus_List" runat="server" CssClass="text-Right-Blue" Text="單據狀態" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eSheetStatus_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                                    <asp:Label ID="eSheetStatus_List" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbStatusDate_List" runat="server" CssClass="text-Right-Blue" Text="到貨日" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eStatusDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="開單日" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="開單人員" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                                    <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbPayMode_List" runat="server" CssClass="text-Right-Blue" Text="付款方式" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="ePayMode_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PayMode_C") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbPayDate_List" runat="server" CssClass="text-Right-Blue" Text="付款日" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="ePayDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PayDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_High ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbSheetNote_List" runat="server" CssClass="text-Right-Blue" Text="摘要" Width="100%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="9">
                                    <asp:TextBox ID="eSheetNote_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("SheetNote") %>' Width="95%" Height="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbAmount_List" runat="server" CssClass="text-Right-Blue" Text="小計金額" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbTaxType_List" runat="server" CssClass="text-Right-Blue" Text="稅別" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eTaxType_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TaxType_C") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbTaxRate_List" runat="server" CssClass="text-Right-Blue" Text="稅率" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eTaxRate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TaxRate") %>' Width="75%" />
                                    <asp:Label ID="lbSplit_5" runat="server" CssClass="text-Left-Black" Text="％" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbTaxAMT_List" runat="server" CssClass="text-Right-Blue" Text="稅額" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eTaxAMT_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TaxAMT") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbTotalAmount_List" runat="server" CssClass="text-Right-Blue" Text="總金額" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-10Col">
                                    <asp:Label ID="eTotalAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TotalAmount") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_High ColBorder ColWidth-10Col">
                                    <asp:Label ID="lbRemarkA_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="100%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="9">
                                    <asp:TextBox ID="eRemarkA_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("RemarkA") %>' Width="95%" Height="97%" />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </ItemTemplate>
        </asp:FormView>
        <asp:GridView ID="gridIASheetB_List" runat="server" CellPadding="4" GridLines="None" AllowPaging="True" AutoGenerateColumns="False" DataKeyNames="SheetNoItems" DataSourceID="dsIASheetB_List" ForeColor="#333333" PageSize="5" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="SheetNoItems" HeaderText="序號" ReadOnly="True" SortExpression="SheetNoItems" Visible="False" />
                <asp:BoundField DataField="SheetNo" HeaderText="單號" SortExpression="SheetNo" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                <asp:BoundField DataField="ConsNo" HeaderText="料號" SortExpression="ConsNo" />
                <asp:BoundField DataField="ConsName" HeaderText="品名" SortExpression="ConsName" />
                <asp:BoundField DataField="Unit" HeaderText="Unit" SortExpression="Unit" Visible="False" />
                <asp:BoundField DataField="Unit_C" HeaderText="單位" ReadOnly="True" SortExpression="Unit_C" />
                <asp:BoundField DataField="Price" HeaderText="單價" SortExpression="Price" />
                <asp:BoundField DataField="Quantity" HeaderText="數量" SortExpression="Quantity" />
                <asp:BoundField DataField="Amount" HeaderText="小計" SortExpression="Amount" />
                <asp:BoundField DataField="RemarkB" HeaderText="備註" SortExpression="RemarkB" />
                <asp:BoundField DataField="QtyMode" HeaderText="QtyMode" SortExpression="QtyMode" Visible="False" />
                <asp:BoundField DataField="ItemStatus" HeaderText="ItemStatus" SortExpression="ItemStatus" Visible="False" />
                <asp:BoundField DataField="ItemStatus_C" HeaderText="處理進度" ReadOnly="True" SortExpression="ItemStatus_C" />
            </Columns>
            <FooterStyle BackColor="#990000" ForeColor="White" Font-Bold="True" />
            <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#FFCC66" ForeColor="#333333" HorizontalAlign="Center" />
            <RowStyle BackColor="#FFFBD6" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="Navy" />
            <SortedAscendingCellStyle BackColor="#FDF5AC" />
            <SortedAscendingHeaderStyle BackColor="#4D0000" />
            <SortedDescendingCellStyle BackColor="#FCF6C0" />
            <SortedDescendingHeaderStyle BackColor="#820000" />
        </asp:GridView>
        <asp:FormView ID="fvIASheetB_Detail" runat="server" DataKeyNames="SheetNoItems" DataSourceID="dsIASheetB_Detail" Width="100%" OnDataBound="fvIASheetBDetail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOKB_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKB_Edit_Click" Text="確定" Width="120px" />
                &nbsp;
                <asp:Button ID="bbCancelB_List" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSheetNoB_Edit" runat="server" CssClass="text-Right-Blue" Text="進貨單號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eSheetNoB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' />
                                </td>
                                <td class="ColWidth-8Col" colspan="6">
                                    <asp:Label ID="eSheetNoItems_Edit" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                                    <asp:Label ID="eQtyMode_Edit" runat="server" Text='<%# Eval("QtyMode") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbItems_Edit" runat="server" CssClass="text-Right-Blue" Text="項次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbConsNo_Edit" runat="server" CssClass="text-Right-Blue" Text="料號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eConsNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eConsNo_Edit_TextChanged" Text='<%# Eval("ConsNo") %>' Width="75%" />
                                    <asp:Button ID="bbGetConsNo_Edit" runat="server" CssClass="button-Blue" Text="..." Width="15%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbConsName_Edit" runat="server" CssClass="text-Right-Blue" Text="品名" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:Label ID="eConsName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPrice_Edit" runat="server" CssClass="text-Right-Blue" Text="單價" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="ePrice_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="ePrice_Edit_TextChanged" Text='<%# Eval("Price") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbQuantity_Edit" runat="server" CssClass="text-Right-Blue" Text="數量" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eQuantity_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="ePrice_Edit_TextChanged" Text='<%# Eval("Quantity") %>' Width="50%" />
                                    <asp:Label ID="eUnit_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Unit_C") %>' Width="40%" />
                                    <asp:Label ID="eUnit_Edit" runat="server" Text='<%# Eval("Unit") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAmountB_Edit" runat="server" CssClass="text-Right-Blue" Text="小計" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eAmountB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbItemStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="狀態" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eItemStatus_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                                    <asp:Label ID="eItemStatus_Edit" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemarkB_Edit" runat="server" CssClass="text-Right-Blue" Text="項次備註" Width="100%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemarkB_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-8Col" colspan="3" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDateB_Edit" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDateB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuManB_Edit" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuManB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManNAmeB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-8Col" colspan="3" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDateB_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDateB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyManB_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eModifyManB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManNameB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="60%" />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOKB_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKB_INS_Click" Text="確定" Width="120px" />
                &nbsp;
                <asp:Button ID="bbCancelB_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSheetNoB_INS" runat="server" CssClass="text-Right-Blue" Text="進貨單號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eSheetNoB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' />
                                </td>
                                <td class="ColWidth-8Col" colspan="6">
                                    <asp:Label ID="eSheetNoItems_INS" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                                    <asp:Label ID="eQtyMode_INS" runat="server" Text='<%# Eval("QtyMode") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbItems_INS" runat="server" CssClass="text-Right-Blue" Text="項次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eItems_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbConsNo_INS" runat="server" CssClass="text-Right-Blue" Text="料號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eConsNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eConsNo_INS_TextChanged" Text='<%# Eval("ConsNo") %>' Width="75%" />
                                    <asp:Button ID="bbGetConsNo_INS" runat="server" CssClass="button-Blue" Text="..." Width="15%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbConsName_INS" runat="server" CssClass="text-Right-Blue" Text="品名" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:Label ID="eConsName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPrice_INS" runat="server" CssClass="text-Right-Blue" Text="單價" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="ePrice_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="ePrice_INS_TextChanged" Text='<%# Eval("Price") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbQuantity_INS" runat="server" CssClass="text-Right-Blue" Text="數量" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eQuantity_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="ePrice_INS_TextChanged" Text='<%# Eval("Quantity") %>' Width="50%" />
                                    <asp:Label ID="eUnit_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Unit_C") %>' Width="40%" />
                                    <asp:Label ID="eUnit_INS" runat="server" Text='<%# Eval("Unit") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAmountB_INS" runat="server" CssClass="text-Right-Blue" Text="小計" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eAmountB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbItemStatus_INS" runat="server" CssClass="text-Right-Blue" Text="狀態" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eItemStatus_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                                    <asp:Label ID="eItemStatus_INS" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemarkB_INS" runat="server" CssClass="text-Right-Blue" Text="項次備註" Width="100%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemarkB_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-8Col" colspan="3" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDateB_INS" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDateB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuManB_INS" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuManB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManNameB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-8Col" colspan="3" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDateB_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDateB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyManB_INS" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eModifyManB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManNameB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="60%" />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CssClass="button-Blue" Text="新增明細" CommandName="New" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNewB_List" runat="server" CssClass="button-Black" Text="新增明細" CommandName="New" Width="120px" />
                &nbsp;
                <asp:Button ID="bbEditB_List" runat="server" CssClass="button-Blue" Text="修改明細" CommandName="Edit" Width="120px" />
                &nbsp;
                <asp:Button ID="bbDeleteB_List" runat="server" CssClass="button-Red" Text="刪除明細" OnClick="bbDeleteB_List_Click" Width="120px" />
                &nbsp;
                <asp:Button ID="bbArriveB_List" runat="server" CssClass="button-Blue" Text="明細到貨" OnClick="bbArriveB_List_Click" Width="120px" />
                &nbsp;
                <asp:Button ID="bbIsAbortB_List" runat="server" CssClass="button-Red" Text="明細作廢" OnClick="bbIsAbortB_List_Click" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSheetNoB_List" runat="server" CssClass="text-Right-Blue" Text="進貨單號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eSheetNoB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' />
                                </td>
                                <td class="ColWidth-8Col" colspan="6">
                                    <asp:Label ID="eSheetNoItems_List" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                                    <asp:Label ID="eQtyMode_List" runat="server" Text='<%# Eval("QtyMode") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbItems_List" runat="server" CssClass="text-Right-Blue" Text="項次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbConsNo_List" runat="server" CssClass="text-Right-Blue" Text="料號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eConsNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbConsName_List" runat="server" CssClass="text-Right-Blue" Text="品名" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:Label ID="eConsName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPrice_List" runat="server" CssClass="text-Right-Blue" Text="單價" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="ePrice_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Price") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbQuantity_List" runat="server" CssClass="text-Right-Blue" Text="數量" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eQuantity_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Quantity") %>' Width="50%" />
                                    <asp:Label ID="eUnit_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Unit_C") %>' Width="45%" />
                                    <asp:Label ID="eUnit_List" runat="server" Text='<%# Eval("Unit") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAmountB_List" runat="server" CssClass="text-Right-Blue" Text="小計" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eAmountB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbItemStatus_List" runat="server" CssClass="text-Right-Blue" Text="狀態" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eItemStatus_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                                    <asp:Label ID="eItemStatus_List" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemarkB_List" runat="server" CssClass="text-Right-Blue" Text="項次備註" Width="100%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemarkB_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("RemarkB") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-8Col" colspan="3" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDateB_List" runat="server" CssClass="text-Right-Blue" Text="開單日期" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDateB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuManB_List" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuManB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManNameB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-8Col" colspan="3" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDateB_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDateB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyManB_List" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eModifyManB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManNameB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="60%" />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="dsSheetStatus_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        SelectCommand="SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材進出單fmIASheetA      SheetStatus')" />
    <asp:SqlDataSource ID="dsPayMode_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        SelectCommand="SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材進出單fmIASheetA      PayMode') ORDER BY ClassNo" />
    <asp:SqlDataSource ID="dsPayMode_Edit" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        SelectCommand="SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材進出單fmIASheetA      PayMode') ORDER BY ClassNo" />
    <asp:SqlDataSource ID="dsPayMode_INS" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        SelectCommand="SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材進出單fmIASheetA      PayMode') ORDER BY ClassNo" />
    <asp:SqlDataSource ID="dsBrand_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        SelectCommand="SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材管理  fmIAConsumables Brand') ORDER BY ClassNo" />
    <asp:SqlDataSource ID="dsIASheetA_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        SelectCommand="SELECT a.SheetNo, a.SupNo, (SELECT name FROM CUSTOM WHERE (types = 'S') AND (code = a.SupNo)) AS SupName, a.BuDate, a.TotalAmount, a.SheetNote, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材進出單fmIASheetA      SheetStatus') AND (CLASSNO = a.SheetStatus)) AS SheetStatus_C, a.RemarkA FROM IASheetA AS a LEFT OUTER JOIN IASheetB AS b ON b.SheetNo = a.SheetNo LEFT OUTER JOIN IAConsumables AS c ON c.ConsNo = b.ConsNo WHERE (ISNULL(a.SheetNo, '') = '')"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsIASheetA_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        SelectCommand="SELECT SheetNo, SheetMode, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材進出單fmIASheetA      SheetMode') AND (CLASSNO = a.SheetMode)) AS SheetMode_C, SupNo, (SELECT name FROM CUSTOM WHERE (types = 'S') AND (code = a.SupNo)) AS SupName, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuManName, SheetNote, Amount, TaxType, (SELECT CLASSTXT FROM DBDICB AS DBDICB_3 WHERE (FKEY = '耗材進貨單A     IASheetA        TAXTYPE') AND (CLASSNO = a.TaxType)) AS TaxType_C, TaxRate, TaxAMT, TotalAmount, SheetStatus, (SELECT CLASSTXT FROM DBDICB AS DBDICB_2 WHERE (FKEY = '電腦課耗材進出單fmIASheetA      SheetStatus') AND (CLASSNO = a.SheetStatus)) AS SheetStatus_C, StatusDate, PayDate, PayMode, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (FKEY = '電腦課耗材進出單fmIASheetA      PayMode') AND (CLASSNO = a.PayMode)) AS PayMode_C, RemarkA FROM IASheetA AS a WHERE (SheetNo = @SheetNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridIASheetList" Name="SheetNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsTaxType_Edit" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '耗材進貨單A     IASheetA        TAXTYPE') ORDER BY ClassNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsTaxType_INS" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '耗材進貨單A     IASheetA        TAXTYPE') ORDER BY ClassNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsIASheetB_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT b.SheetNoItems, b.SheetNo, b.Items, b.ConsNo, c.ConsName, c.Unit, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材管理  fmIAConsumables Unit') AND (CLASSNO = c.Unit)) AS Unit_C, b.Price, b.Quantity, b.Amount, b.RemarkB, b.QtyMode, b.ItemStatus, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (FKEY = '電腦課耗材進出單fmIASheetB      ItemStatus') AND (CLASSNO = b.ItemStatus)) AS ItemStatus_C FROM IASheetB AS b LEFT OUTER JOIN IAConsumables AS c ON c.ConsNo = b.ConsNo WHERE (b.SheetNo = @SheetNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridIASheetList" Name="SheetNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsIASheetB_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT b.SheetNoItems, b.SheetNo, b.Items, b.ConsNo, c.ConsName, c.Unit, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材管理  fmIAConsumables Unit') AND (CLASSNO = c.Unit)) AS Unit_C, b.Price, b.Quantity, b.Amount, b.RemarkB, b.QtyMode, b.ItemStatus, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (FKEY = '電腦課耗材進出單fmIASheetB      ItemStatus') AND (CLASSNO = b.ItemStatus)) AS ItemStatus_C, b.BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = b.BuMan)) AS BuMan_C, b.BuDate, b.ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = b.ModifyMan)) AS ModifyMan_C, b.ModifyDate FROM IASheetB AS b LEFT OUTER JOIN IAConsumables AS c ON c.ConsNo = b.ConsNo WHERE (b.SheetNoItems = @SheetNoItems)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridIASheetB_List" Name="SheetNoItems" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
</asp:Content>
