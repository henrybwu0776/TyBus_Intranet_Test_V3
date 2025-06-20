<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ConsumablesInstore.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ConsumablesInstore" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
<asp:Content ID="ConsumablesInstoreForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">總務耗材進貨單</a>
    </div>
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Search_TextChanged" Width="30%" />
                    <asp:Label ID="eDepName_Search" runat="server" CssClass="text-Left-Black" Width="65%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbAssignMan_Search" runat="server" CssClass="text-Right-Blue" Text="申請人" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eAssignMan_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignMan_Search_TextChanged" Width="30%" />
                    <asp:Label ID="eAssignManName_Search" runat="server" CssClass="text-Left-Black" Width="65%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbBuDate_Search" runat="server" CssClass="text-Right-Blue" Text="申請日期" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eBuDate_S_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eBuDate_E_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-9Col" colspan="9">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Blue" Text="查詢" OnClick="bbSearch_Click" Width="120px" />
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" OnClick="bbExit_Click" Width="120px" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData_A" runat="server" CssClass="ShowPanel">
        <div>
            <asp:Label ID="eErrorMSG_A" runat="server" CssClass="errorMessageText" Visible="false" />
        </div>
        <asp:GridView ID="gridConsSheetA_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" CellPadding="3" DataKeyNames="SheetNo" DataSourceID="sdsConsSheetA_List" GridLines="Horizontal" PageSize="5" Width="100%">
            <AlternatingRowStyle BackColor="#F7F7F7" />
            <Columns>
                <asp:CommandField ButtonType="Button" InsertVisible="False" ShowCancelButton="False" ShowSelectButton="True" />
                <asp:BoundField DataField="SheetNo" HeaderText="進貨單號" ReadOnly="True" SortExpression="SheetNo" />
                <asp:BoundField DataField="SupNo" HeaderText="SupNo" SortExpression="SupNo" Visible="False" />
                <asp:BoundField DataField="SupName" HeaderText="廠商" SortExpression="SupName" />
                <asp:BoundField DataField="SheetStatus" HeaderText="SheetStatus" SortExpression="SheetStatus" Visible="False" />
                <asp:BoundField DataField="SheetStatus_C" HeaderText="單據狀態" SortExpression="SheetStatus_C" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="申請站別" SortExpression="DepName" />
                <asp:BoundField DataField="TotalAmount" HeaderText="單據總金額" SortExpression="TotalAmount" />
                <asp:BoundField DataField="BuDate" HeaderText="申請日期" ReadOnly="True" SortExpression="BuDate" />
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
        <asp:SqlDataSource ID="sdsConsSheetA_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select a.SheetNo, a.SupNo, c.[Name] SupName, a.SheetStatus, d.ClassTxt SheetStatus_C,
       a.DepNo, b.[Name] DepName, TotalAmount, convert(varchar(10), a.BuDate, 111) BuDate
  from ConsSheetA a left join [Custom] c on c.Code = a.SupNo and c.[Types] = 'S'
                    left join DBDICB d on d.ClassNo = a.SheetStatus and d.FKey = '總務耗材進出單  fmConsSheetA    SheetStatus'
					left join Department b on b.DepNo = a.DepNo
 where a.SheetMode = 'SI'"></asp:SqlDataSource>
        <asp:FormView ID="fvConsSheetA_Detail" runat="server" DataKeyNames="SheetNo" DataSourceID="sdsConsSheetA_Detail" OnDataBound="fvConsSheetA_Detail_DataBound" Width="100%">
            <EditItemTemplate>
                <asp:Button ID="bbOK_A_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_A_Edit_Click" Text="確定" Width="120px" />
                <asp:Button ID="bbCancel_A_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetNoA_Edit" runat="server" CssClass="text-Right-Blue" Text="進貨單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetNoA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateA_Edit" runat="server" CssClass="text-Right-Blue" Text="進貨日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuDateA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Width="35%" />
                            <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAssignMan_Edit" runat="server" CssClass="text-Right-Blue" Text="申請人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eAssignMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' AutoPostBack="true" OnTextChanged="eAssignMan_Edit_TextChanged" Width="35%" />
                            <asp:Label ID="eAssignManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder  ColWidth-10Col">
                            <asp:Label ID="lbAmountA_Edit" runat="server" CssClass="text-Right-Blue" Text="金額小計" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAmountA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lnTaxType_Edit" runat="server" CssClass="text-Right-Blue" Text="稅別" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:DropDownList ID="eTaxType_C_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="eTaxType_C_Edit_SelectedIndexChanged" Width="95%" />
                            <asp:Label ID="eTaxType_Edit" runat="server" Text='<%# Eval("TaxType") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTaxRate_Edit" runat="server" CssClass="text-Right-Blue" Text="稅率" Width="95%" />
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
                            <asp:Label ID="lbSupNo_Edit" runat="server" CssClass="text-Right-Blue" Text="供應商" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                            <asp:TextBox ID="eSupNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupNo") %>' Width="25%" />
                            <asp:TextBox ID="eSupName_Edit" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("SupName") %>' Width="70%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3">
                            <asp:Label ID="lbRemarkA_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3" colspan="5">
                            <asp:TextBox ID="eRemarkA_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPayMode_Edit" runat="server" CssClass="text-Right-Blue" Text="付款方式" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:DropDownList ID="ePayMode_C_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ePayMode_C_Edit_SelectedIndexChanged" Width="95%" />
                            <asp:Label ID="ePayMode_Edit" runat="server" Text='<%# Eval("PayMode") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPayDate_Edit" runat="server" CssClass="text-Right-Blue" Text="付款日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePayDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PayDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="單據狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetStatus_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                            <asp:Label ID="eSheetStatus_Edit" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbStatusDate_Edit" runat="server" CssClass="text-Right-Blue" Text="狀態變更日" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eStatusDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuManA_Edit" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManNameA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateA_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManA_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManNameA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
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
                            <asp:Label ID="lbSheetNoA_INS" runat="server" CssClass="text-Right-Blue" Text="進貨單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetNoA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateA_INS" runat="server" CssClass="text-Right-Blue" Text="進貨日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuDateA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Width="35%" />
                            <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAssignMan_INS" runat="server" CssClass="text-Right-Blue" Text="申請人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eAssignMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' AutoPostBack="true" OnTextChanged="eAssignMan_INS_TextChanged" Width="35%" />
                            <asp:Label ID="eAssignManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder  ColWidth-10Col">
                            <asp:Label ID="lbAmountA_INS" runat="server" CssClass="text-Right-Blue" Text="金額小計" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAmountA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lnTaxType_INS" runat="server" CssClass="text-Right-Blue" Text="稅別" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:DropDownList ID="eTaxType_C_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="eTaxType_C_INS_SelectedIndexChanged" Width="95%" />
                            <asp:Label ID="eTaxType_INS" runat="server" Text='<%# Eval("TaxType") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTaxRate_INS" runat="server" CssClass="text-Right-Blue" Text="稅率" Width="95%" />
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
                            <asp:Label ID="lbSupNo_INS" runat="server" CssClass="text-Right-Blue" Text="供應商" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                            <asp:TextBox ID="eSupNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupNo") %>' Width="25%" />
                            <asp:TextBox ID="eSupName_INS" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("SupName") %>' Width="70%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3">
                            <asp:Label ID="lbRemarkA_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3" colspan="5">
                            <asp:TextBox ID="eRemarkA_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPayMode_INS" runat="server" CssClass="text-Right-Blue" Text="付款方式" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:DropDownList ID="ePayMode_C_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ePayMode_C_INS_SelectedIndexChanged" Width="95%" />
                            <asp:Label ID="ePayMode_INS" runat="server" Text='<%# Eval("PayMode") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPayDate_INS" runat="server" CssClass="text-Right-Blue" Text="付款日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePayDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PayDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetStatus_INS" runat="server" CssClass="text-Right-Blue" Text="單據狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetStatus_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                            <asp:Label ID="eSheetStatus_INS" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbStatusDate_INS" runat="server" CssClass="text-Right-Blue" Text="狀態變更日" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eStatusDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuManA_INS" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManNameA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateA_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManA_INS" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManNameA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
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
                <asp:Button ID="bbNewA_Empty" runat="server" CssClass="button-Blue" Text="新增" CommandName="New" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNewA_List" runat="server" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
                <asp:Button ID="bbEditA_List" runat="server" CssClass="button-Blue" CommandName="Edit" Text="修改" Width="120px" />
                <asp:Button ID="bbPrint_List" runat="server" CssClass="button-Black" Text="列印進貨單" OnClick="bbPrint_List_Click" Width="120px" />
                <asp:Button ID="bbCloseA_List" runat="server" CssClass="button-Blue" Text="全部入庫" OnClick="bbCloseA_List_Click" Width="120px" />
                <asp:Button ID="bbDeleteA_List" runat="server" CssClass="button-Red" Text="刪除" OnClick="bbDeleteA_List_Click" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetNoA_List" runat="server" CssClass="text-Right-Blue" Text="進貨單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetNoA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateA_List" runat="server" CssClass="text-Right-Blue" Text="進貨日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuDateA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAssignMan_List" runat="server" CssClass="text-Right-Blue" Text="申請人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eAssignMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' Width="35%" />
                            <asp:Label ID="eAssignManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder  ColWidth-10Col">
                            <asp:Label ID="lbAmountA_List" runat="server" CssClass="text-Right-Blue" Text="金額小計" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAmountA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lnTaxType_List" runat="server" CssClass="text-Right-Blue" Text="稅別" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eTaxType_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TaxType_C") %>' Width="95%" />
                            <asp:Label ID="eTaxType_LIst" runat="server" Text='<%# Eval("TaxType") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTaxRate_List" runat="server" CssClass="text-Right-Blue" Text="稅率" Width="95%" />
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
                            <asp:Label ID="lbSupNo_List" runat="server" CssClass="text-Right-Blue" Text="供應商" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                            <asp:Label ID="eSupNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupNo") %>' Width="25%" />
                            <asp:Label ID="eSupName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupName") %>' Width="70%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3">
                            <asp:Label ID="lbRemarkA_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3" colspan="5">
                            <asp:TextBox ID="eRemarkA_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("RemarkA") %>' Height="95%" Width="97%" />
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
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetStatus_List" runat="server" CssClass="text-Right-Blue" Text="單據狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetStatus_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                            <asp:Label ID="eSheetStatus_List" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbStatusDate_List" runat="server" CssClass="text-Right-Blue" Text="狀態變更日" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eStatusDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuManA_List" runat="server" CssClass="text-Right-Blue" Text="開單人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManNameA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateA_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManA_List" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManNameA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
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
        <asp:SqlDataSource ID="sdsConsSheetA_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select a.SheetNo, a.SupNo, c.[Name] SupName, convert(varchar(10), a.BuDate, 111) BuDate, a.AssignMan, e.[Name] AssignManName,
       a.BuMan, e2.[Name] BuManName, a.Amount, a.TaxType, d2.ClassTxt TaxType_C, a.TaxRate, a.TaxAMT, a.TotalAmount, 
	   a.SheetStatus, d.ClassTxt SheetStatus_C, convert(varchar(10), a.StatusDate, 111) StatusDate, 
	   convert(varchar(10), a.PayDate, 111) PayDate, a.PayMode, d3.ClassTxt PayMode_C, a.RemarkA, a.DepNo, b.[Name] DepName,
	   a.ModifyMan, e3.[Name] ModifyManName, convert(varchar(10), a.ModifyDate, 111) ModifyDate
  from ConsSheetA a left join Department b on b.DepNo = a.DepNo
                    left join [Custom] c on c.Code = a.SupNo and c.[Types] = 'S'
                    left join DBDICB d on d.ClassNo = a.SheetStatus and d.FKey = '總務耗材進出單  fmConsSheetA    SheetStatus'
					left join DBDICB d2 on d2.ClassNo = a.TaxType and d2.FKey = '總務請購單      ConsSheetA      TAXTYPE'
					left join DBDICB d3 on d3.ClassNo = a.PayMode and d3.FKey = '總務請購單      ConsSheetA      PayMode'
					left join Employee e on e.EmpNo = a.AssignMan
					left join Employee e2 on e2.EmpNo = a.BuMan
					left join Employee e3 on e3.EmpNo = a.ModifyMan
 where a.SheetNo = @SheetNo">
            <SelectParameters>
                <asp:ControlParameter ControlID="gridConsSheetA_List" Name="SheetNo" PropertyName="SelectedValue" />
            </SelectParameters>
        </asp:SqlDataSource>
    </asp:Panel>
    <asp:Panel ID="plShowData_B" runat="server" CssClass="ShowPanel-Detail">
        <div>
            <asp:Label ID="eErrorMSG_B" runat="server" CssClass="errorMessageText" Visible="false" />
        </div>
        <asp:GridView ID="gridConsSheetB_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="SheetNoItems" DataSourceID="sdsConsSheetB_List" ForeColor="#333333" GridLines="None" PageSize="5" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" InsertVisible="False" ShowCancelButton="False" ShowSelectButton="True" />
                <asp:BoundField DataField="SheetNoItems" HeaderText="SheetNoItems" ReadOnly="True" SortExpression="SheetNoItems" Visible="False" />
                <asp:BoundField DataField="SheetNo" HeaderText="SheetNo" SortExpression="SheetNo" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                <asp:BoundField DataField="ConsNo" HeaderText="料號" SortExpression="ConsNo" />
                <asp:BoundField DataField="ConsName" HeaderText="品名" SortExpression="ConsName" />
                <asp:BoundField DataField="Price" HeaderText="單價" SortExpression="Price" />
                <asp:BoundField DataField="Quantity" HeaderText="數量" SortExpression="Quantity" />
                <asp:BoundField DataField="ConsUnit" HeaderText="ConsUnit" SortExpression="ConsUnit" Visible="False" />
                <asp:BoundField DataField="ConsUnit_C" HeaderText="單位" SortExpression="ConsUnit_C" />
                <asp:BoundField DataField="ItemStatus" HeaderText="ItemStatus" SortExpression="ItemStatus" Visible="False" />
                <asp:BoundField DataField="ItemStatus_C" HeaderText="明細狀態" SortExpression="ItemStatus_C" />
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
        <asp:SqlDataSource ID="sdsConsSheetB_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select b.SheetNoItems, b.SheetNo, b.Items, b.ConsNo, c.ConsName, b.Price, b.Quantity, b.ConsUnit, d2.ClassTxt ConsUnit_C, b.ItemStatus, d1.ClassTxt ItemStatus_C
  from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo
                    left join DBDICB d1 on d1.ClassNo = b.ItemStatus and d1.FKey = '總務課耗材進出單ConsSheetB      ItemStatus'
					left join DBDICB d2 on d2.ClassNo = b.ConsUnit and d2.FKey = '耗材庫存        CONSUMABLES     ConsUnit'
where b.SheetNo = @SheetNo
order by b.Items">
            <SelectParameters>
                <asp:ControlParameter ControlID="gridConsSheetA_List" Name="SheetNo" PropertyName="SelectedValue" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:FormView ID="fvConsSheetB_Detail" runat="server" DataKeyNames="SheetNoItems" DataSourceID="sdsConsSheetB_Detail" Width="100%" OnDataBound="fvConsSheetB_Detail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOKB_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKB_Edit_Click" Text="確定" Width="120px" />
                <asp:Button ID="bbCancelB_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItems_Edit" runat="server" CssClass="text-Right-Blue" Text="明細項次" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoB_Edit" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                            <asp:Label ID="eSheetNoItems_Edit" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbConsNo_Edit" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="5">
                            <asp:TextBox ID="eConsNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="35%" />
                            <asp:TextBox ID="eConsName_Edit" runat="server" CssClass="text-Left-Black" Enabled="false" AutoPostBack="true" OnTextChanged="eConsName_Edit_TextChanged" Text='<%# Eval("ConsName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItemStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="明細狀況" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItemStatus_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lnStockQty_Edit" runat="server" CssClass="text-Right-Blue" Text="現有庫存量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eStockQty_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StockQty") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbQuantity_Edit" runat="server" CssClass="text-Right-Blue" Text="數量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eQuantity_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eQuantity_Edit_TextChanged" Text='<%# Eval("Quantity") %>' Width="55%" />
                            <asp:Label ID="eConsUnit_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="40%" />
                            <asp:Label ID="eConsUnit_Edit" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPrice_Edit" runat="server" CssClass="text-Right-Blue" Text="單價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePrice_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eQuantity_Edit_TextChanged" Text='<%# Eval("Price") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAVGPrice_Edit" runat="server" CssClass="text-Right-Blue" Text="平均單價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAVGPrice_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AVGPrice") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAmountB_Edit" runat="server" CssClass="text-Right-Blue" Text="小計" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAmountB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColBorder ColWidth-10Col MultiLine_High">
                            <asp:Label ID="lbRemarkB_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                            <asp:TextBox ID="eRemarkB_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
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
                            <asp:Label ID="lbBuManB_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManNameB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateB_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManNameB_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManNameB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
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
                            <asp:Label ID="lbItems_INS" runat="server" CssClass="text-Right-Blue" Text="明細項次" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItems_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoB_INS" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                            <asp:Label ID="eSheetNoItems_INS" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbConsNo_INS" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="5">
                            <asp:TextBox ID="eConsNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="35%" />
                            <asp:TextBox ID="eConsName_INS" runat="server" CssClass="text-Left-Black" Enabled="false" AutoPostBack="true" OnTextChanged="eConsName_INS_TextChanged" Text='<%# Eval("ConsName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItemStatus_INS" runat="server" CssClass="text-Right-Blue" Text="明細狀況" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItemStatus_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lnStockQty_INS" runat="server" CssClass="text-Right-Blue" Text="現有庫存量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eStockQty_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StockQty") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbQuantity_INS" runat="server" CssClass="text-Right-Blue" Text="數量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eQuantity_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eQuantity_INS_TextChanged" Text='<%# Eval("Quantity") %>' Width="55%" />
                            <asp:Label ID="eConsUnit_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="40%" />
                            <asp:Label ID="eConsUnit_INS" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPrice_INS" runat="server" CssClass="text-Right-Blue" Text="單價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePrice_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eQuantity_INS_TextChanged" Text='<%# Eval("Price") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAVGPrice_INS" runat="server" CssClass="text-Right-Blue" Text="平均單價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAVGPrice_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AVGPrice") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAmountB_INS" runat="server" CssClass="text-Right-Blue" Text="小計" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAmountB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColBorder ColWidth-10Col MultiLine_High">
                            <asp:Label ID="lbRemarkB_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                            <asp:TextBox ID="eRemarkB_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
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
                            <asp:Label ID="lbBuManB_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManNameB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateB_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManNameB_INS" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManNameB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
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
                <asp:Button ID="bbNewB_Empty" runat="server" CssClass="button-Blue" Text="新增明細" CommandName="New" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNewB_List" runat="server" CssClass="button-Black" Text="新增明細" CommandName="New" Width="120px" />
                <asp:Button ID="bbEditB_List" runat="server" CssClass="button-Blue" Text="修改明細" CommandName="Edit" Width="120px" />
                <asp:Button ID="bbArriveB_List" runat="server" CssClass="button-Black" Text="明細到貨" Width="120px" />
                <asp:Button ID="bbDeleteB_List" runat="server" CssClass="button-Red" Text="刪除明細" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItems_List" runat="server" CssClass="text-Right-Blue" Text="明細項次" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoB_List" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                            <asp:Label ID="eSheetNoItems_List" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbConsNo_List" runat="server" CssClass="text-Right-Blue" Text="品名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="5">
                            <asp:Label ID="eConsNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="35%" />
                            <asp:Label ID="eConsName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbItemStatus_List" runat="server" CssClass="text-Right-Blue" Text="明細狀況" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItemStatus_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lnStockQty_List" runat="server" CssClass="text-Right-Blue" Text="現有庫存量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eStockQty_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StockQty") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbQuantity_List" runat="server" CssClass="text-Right-Blue" Text="數量" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eQuantity_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Quantity") %>' Width="55%" />
                            <asp:Label ID="eConsUnit_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="40%" />
                            <asp:Label ID="eConsUnit_List" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPrice_List" runat="server" CssClass="text-Right-Blue" Text="單價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePrice_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Price") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAVGPrice_List" runat="server" CssClass="text-Right-Blue" Text="平均單價" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAVGPrice_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AVGPrice") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAmountB_List" runat="server" CssClass="text-Right-Blue" Text="小計" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAmountB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColBorder ColWidth-10Col MultiLine_High">
                            <asp:Label ID="lbRemarkB_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                            <asp:TextBox ID="eRemarkB_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
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
                            <asp:Label ID="lbBuManB_List" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuManB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManNameB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateB_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManNameB_List" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyManB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManNameB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
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
        <asp:SqlDataSource ID="sdsConsSheetB_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select b.SheetNoItems, b.SheetNo, b.Items, b.ConsNo, c.ConsName, b.Price, b.Quantity, b.ConsUnit, d2.ClassTxt ConsUnit_C, b.ItemStatus, d1.ClassTxt ItemStatus_C,
       b.Amount, c.StockQty, c.AVGPrice, b.RemarkB, b.BuMan, e1.[Name] BuManName, b.BuDate, b.ModifyMan, e2.[Name] ModifyManName, b.ModifyDate
  from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo
                    left join DBDICB d1 on d1.ClassNo = b.ItemStatus and d1.FKey = '總務課耗材進出單ConsSheetB      ItemStatus'
					left join DBDICB d2 on d2.ClassNo = b.ConsUnit and d2.FKey = '耗材庫存        CONSUMABLES     ConsUnit'
					left join Employee e1 on e1.EmpNo = b.BuMan
					left join Employee e2 on e2.EmpNo = b.ModifyMan
where b.SheetNoItems = @SheetNoItems">
            <SelectParameters>
                <asp:ControlParameter ControlID="gridConsSheetB_List" Name="SheetNoItems" PropertyName="SelectedValue" />
            </SelectParameters>
        </asp:SqlDataSource>
    </asp:Panel>
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
        <asp:Button ID="ClosePrint" runat="server" CssClass="button-Red" Text="結束預覽" Width="120px" />
        <br />
        <rsweb:ReportViewer ID="rvPrint" runat="server" Width="100%"></rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
