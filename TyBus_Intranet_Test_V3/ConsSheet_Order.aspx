<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ConsSheet_Order.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ConsSheet_Order" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="ConsSheet_OrderForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">總務耗材請購單</a>
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
    <asp:Panel ID="plShowData_A" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridConsSheetA_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="SheetNo" DataSourceID="sdsConsSheetA_List" ForeColor="#333333" GridLines="None" PageSize="5" Width="100%" OnPageIndexChanging="gridConsSheetA_List_PageIndexChanging">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" InsertVisible="False" ShowCancelButton="False" ShowSelectButton="True" />
                <asp:BoundField DataField="SheetNo" HeaderText="單號" ReadOnly="True" SortExpression="SheetNo" />
                <asp:BoundField DataField="SheetNote" HeaderText="主旨" ReadOnly="true" SortExpression="SheetNote" />
                <asp:BoundField DataField="DepNo" HeaderText="單位代碼" SortExpression="DepNo" Visible="false" />
                <asp:BoundField DataField="DepName" HeaderText="單位" SortExpression="DepName" />
                <asp:BoundField DataField="AssignMan" HeaderText="申請人" SortExpression="AssignMan" Visible="false" />
                <asp:BoundField DataField="AssignManName" HeaderText="申請人姓名" SortExpression="AssignManName" />
                <asp:BoundField DataField="BuDate" DataFormatString="{0:yyyy/MM/dd}" HeaderText="申請日期" ReadOnly="True" SortExpression="BuDate" />
                <asp:BoundField DataField="TotalAmount" HeaderText="總金額" SortExpression="TotalAmount" />
                <asp:BoundField DataField="SheetStatus_C" HeaderText="目前進度" ReadOnly="True" SortExpression="SheetStatus_C" />
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
    </asp:Panel>
    <asp:SqlDataSource ID="sdsConsSheetA_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT a.SheetNo, a.SheetNote, a.DepNo, d.NAME AS DepName, a.AssignMan, e.NAME AS AssignManName, CONVERT (varchar(10), a.BuDate, 111) AS BuDate, a.TotalAmount, CASE WHEN a.SheetStatus = '0' THEN '已開單' WHEN a.SheetStatus = '1' THEN '處理中' WHEN a.SheetStatus = '8' THEN '已作廢' WHEN a.SheetStatus = '9' THEN '已結案' END AS SheetStatus_C FROM ConsSheetA AS a LEFT OUTER JOIN DEPARTMENT AS d ON d.DEPNO = a.DepNo LEFT OUTER JOIN EMPLOYEE AS e ON e.EMPNO = a.AssignMan
where isnull(a.SheetMode, '') = 'BS' "></asp:SqlDataSource>
    <div>
        <asp:Label ID="eErrorMSG_A" runat="server" CssClass="errorMessageText" Visible="false" />
    </div>
    <asp:FormView ID="fvConsSheetA_Detail" runat="server" DataKeyNames="SheetNo" DataSourceID="sdsConsSheetA_Detail" OnDataBound="fvConsSheetA_Detail_DataBound" Width="100%">
        <EditItemTemplate>
            <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_Edit_Click" Text="確定" Width="120px" />
            <asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
            <table class="TableSetting">
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbSheetNo_Edit" runat="server" CssClass="text-Right-Blue" Text="請購單號" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eSheetNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbSupNo_Edit" runat="server" CssClass="text-Right-Blue" Text="供應商" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="7">
                        <asp:TextBox ID="eSupNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupNo") %>' Width="25%" />
                        <asp:TextBox ID="eSupName_Edit" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("SupName") %>' Width="65%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="申請日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Width="35%" />
                        <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbAssignMan_Edit" runat="server" CssClass="text-Right-Blue" Text="申請人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:TextBox ID="eAssignMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' AutoPostBack="true" OnTextChanged="eAssignMan_Edit_TextChanged" Width="35%" />
                        <asp:Label ID="eAssignManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="60%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="金額小計" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eAmount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbTaxType_Edit" runat="server" CssClass="text-Right-Blue" Text="稅別" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:DropDownList ID="eTaxType_C_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="eTaxType_C_Edit_SelectedIndexChanged" Width="95%" />
                        <asp:Label ID="eTaxType_Edit" runat="server" Text='<%# Eval("TaxType") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbTaxRate_Edit" runat="server" CssClass="text-Left-Black" Text="稅率" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eTaxRate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TaxRate") %>' AutoPostBack="true" OnTextChanged="eTaxType_C_Edit_SelectedIndexChanged" Width="95%" />
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
                        <asp:Label ID="lbSheetStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="單據狀態" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eSheetStatus_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                        <asp:Label ID="eSheetStatus_Edit" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbSheetNote_Edit" runat="server" CssClass="text-Right-Blue" Text="主旨" Width="95%" />
                    </td>
                    <td class="ColBorder ColWidth-10Col MultiLine_Low" colspan="7">
                        <asp:TextBox ID="eSheetNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("SheetNote") %>' Height="97%" Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbStatusDate_Edit" runat="server" CssClass="text-Right-Blue" Text="狀態日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eStatusDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3">
                        <asp:Label ID="lbRemarkA_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3" colspan="7">
                        <asp:TextBox ID="eRemarkA_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Height="97%" Width="97%" />
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
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbPayDate_Edit" runat="server" CssClass="text-Right-Blue" Text="付款日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="ePayDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PayDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lnBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                        <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                        <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
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
            <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_INS_Click" Text="確定" Width="120px" />
            <asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
            <table class="TableSetting">
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbSheetNo_INS" runat="server" CssClass="text-Right-Blue" Text="請購單號" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eSheetNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbSupNo_INS" runat="server" CssClass="text-Right-Blue" Text="供應商" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="7">
                        <asp:TextBox ID="eSupNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupNo") %>' Width="25%" />
                        <asp:TextBox ID="eSupName_INS" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("SupName") %>' Width="65%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="申請日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Width="35%" />
                        <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbAssignMan_INS" runat="server" CssClass="text-Right-Blue" Text="申請人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:TextBox ID="eAssignMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' AutoPostBack="true" OnTextChanged="eAssignMan_INS_TextChanged" Width="35%" />
                        <asp:Label ID="eAssignManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="60%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbAmount_INS" runat="server" CssClass="text-Right-Blue" Text="金額小計" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eAmount_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbTaxType_INS" runat="server" CssClass="text-Right-Blue" Text="稅別" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:DropDownList ID="eTaxType_C_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="eTaxType_C_INS_SelectedIndexChanged" Width="95%" />
                        <asp:Label ID="eTaxType_INS" runat="server" Text='<%# Eval("TaxType") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbTaxRate_INS" runat="server" CssClass="text-Left-Black" Text="稅率" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eTaxRate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TaxRate") %>' Width="95%" />
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
                        <asp:Label ID="lbSheetStatus_INS" runat="server" CssClass="text-Right-Blue" Text="單據狀態" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eSheetStatus_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                        <asp:Label ID="eSheetStatus_INS" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbSheetNote_INS" runat="server" CssClass="text-Right-Blue" Text="主旨" Width="95%" />
                    </td>
                    <td class="ColBorder ColWidth-10Col MultiLine_Low" colspan="7">
                        <asp:TextBox ID="eSheetNote_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("SheetNote") %>' Height="97%" Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbStatusDate_INS" runat="server" CssClass="text-Right-Blue" Text="狀態日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eStatusDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3">
                        <asp:Label ID="lbRemarkA_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3" colspan="7">
                        <asp:TextBox ID="eRemarkA_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Height="97%" Width="97%" />
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
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbPayDate_INS" runat="server" CssClass="text-Right-Blue" Text="付款日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="ePayDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PayDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lnBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                        <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                        <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
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
            <asp:Button ID="bbNew_Empty" runat="server" CssClass="button-Blue" Text="新增" CommandName="New" Width="120px" />
        </EmptyDataTemplate>
        <ItemTemplate>
            <asp:Button ID="bbNew_List" runat="server" CssClass="button-Black" Text="新增" CommandName="New" Width="120px" />
            <asp:Button ID="bbEdit_List" runat="server" CssClass="button-Blue" Text="修改" CommandName="Edit" Width="120px" />
            <asp:Button ID="bbPrint_List" runat="server" CssClass="button-Black" Text="列印請購單" OnClick="bbPrint_List_Click" Width="120px" />
            <asp:Button ID="bbAbortA_List" runat="server" CssClass="button-Blue" Text="請購單作廢" OnClick="bbAbortA_List_Click" Width="120px" />
            <asp:Button ID="bbDelete_List" runat="server" CssClass="button-Red" Text="刪除" OnClick="bbDelete_List_Click" Width="120px" />
            <table class="TableSetting">
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbSheetNo_List" runat="server" CssClass="text-Right-Blue" Text="請購單號" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eSheetNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbSupNo_List" runat="server" CssClass="text-Right-Blue" Text="供應商" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="7">
                        <asp:Label ID="eSupNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupNo") %>' Width="25%" />
                        <asp:Label ID="eSupName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SupName") %>' Width="65%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="申請日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                        <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbAssignMan_List" runat="server" CssClass="text-Right-Blue" Text="申請人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:Label ID="eAssignMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' Width="35%" />
                        <asp:Label ID="eAssignManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="60%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbAmount_List" runat="server" CssClass="text-Right-Blue" Text="金額小計" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbTaxType_List" runat="server" CssClass="text-Right-Blue" Text="稅別" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eTaxType_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TaxType_C") %>' Width="95%" />
                        <asp:Label ID="eTaxType_List" runat="server" Text='<%# Eval("TaxType") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbTaxRate_List" runat="server" CssClass="text-Left-Black" Text="稅率" Width="95%" />
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
                        <asp:Label ID="lbSheetStatus_List" runat="server" CssClass="text-Right-Blue" Text="單據狀態" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eSheetStatus_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetStatus_C") %>' Width="95%" />
                        <asp:Label ID="eSheetStatus_List" runat="server" Text='<%# Eval("SheetStatus") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbSheetNote_List" runat="server" CssClass="text-Right-Blue" Text="主旨" Width="95%" />
                    </td>
                    <td class="ColBorder ColWidth-10Col MultiLine_Low" colspan="7">
                        <asp:TextBox ID="eSheetNote_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("SheetNote") %>' Enabled="false" Height="97%" Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbStatusDate_List" runat="server" CssClass="text-Right-Blue" Text="狀態日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eStatusDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StatusDate","{0:yyyy/MM/dd}") %>' />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3">
                        <asp:Label ID="lbRemarkA_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3" colspan="7">
                        <asp:TextBox ID="eRemarkA_List" runat="server" CssClass="text-Left-Black" Enabled="false" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Height="97%" Width="97%" />
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
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbPayDate_List" runat="server" CssClass="text-Right-Blue" Text="付款日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="ePayDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PayDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lnBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                        <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                        <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
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
    <asp:SqlDataSource ID="sdsConsSheetA_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT a.SheetNo, 
       a.SupNo, b.[Name] as SupName, CONVERT (varchar(10), a.BuDate, 111) AS BuDate, a.BuMan,
	   e2.[Name] as BuManName, a.Amount, a.TaxType, z.ClassTxt as TaxType_C, a.TaxRate, a.TaxAMT, 
	   a.TotalAmount, a.SheetStatus, a.SheetNote, z3.ClassTxt as SheetStatus_C, 
	   convert(varchar(10), a.StatusDate, 111) as StatusDate, a.PayDate, a.PayMode, z2.ClassTxt as PayMode_C, 
	   a.RemarkA, a.ModifyMan, e3.[Name] as ModifyManName,
	   convert(varchar(10), a.ModifyDate, 111) as ModifyDate, a.DepNo, d.[Name] as DepName, a.AssignMan,
	   e.[Name] as AssignManName
  FROM ConsSheetA a LEFT OUTER JOIN DEPARTMENT AS d ON d.DEPNO = a.DepNo 
                    LEFT OUTER JOIN EMPLOYEE AS e ON e.EMPNO = a.AssignMan
                    LEFT OUTER JOIN EMPLOYEE AS e2 ON e2.EMPNO = a.BuMan
                    LEFT OUTER JOIN EMPLOYEE AS e3 ON e3.EMPNO = a.ModifyMan
					LEFT OUTER JOIN [Custom] b on b.Code = a.SupNo and b.[Types] = 'S'
					LEFT OUTER JOIN DBDICB as z on z.ClassNo = a.TaxType and z.FKey = '總務請購單      ConsSheetA      TAXTYPE'
					LEFT OUTER JOIN DBDICB as z2 on z2.ClassNo = a.PayMode and z2.FKey = '總務請購單      ConsSheetA      PayMode'
                    LEFT OUTER JOIN DBDICB as z3 on z3.ClassNo = a.SheetStatus and z3.FKey = '總務耗材進出單  fmConsSheetA    SheetStatus'
 WHERE a.SheetNo = @SheetNo">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridConsSheetA_List" Name="SheetNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:Panel ID="plShowDetail" runat="server" CssClass="ShowPanel-Detail">
        <div>
            <asp:Label ID="eErrorMSG_B" runat="server" CssClass="errorMessageText" />
        </div>
        <asp:GridView ID="gridConsSheetB_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="SheetNoItems" DataSourceID="sdsConsSheetB_List" ForeColor="#333333" GridLines="None" PageSize="5" Width="100%" OnPageIndexChanging="gridConsSheetB_List_PageIndexChanging">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="SheetNo" HeaderText="單號" SortExpression="SheetNo" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                <asp:BoundField DataField="SheetNoItems" HeaderText="序號" ReadOnly="True" SortExpression="SheetNoItems" Visible="False" />
                <asp:BoundField DataField="ConsNo" HeaderText="料號" SortExpression="ConsNo" />
                <asp:BoundField DataField="ConsName" HeaderText="品名" SortExpression="ConsName" />
                <asp:BoundField DataField="Price" HeaderText="單價" SortExpression="Price" />
                <asp:BoundField DataField="Quantity" HeaderText="數量" SortExpression="Quantity" />
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
    </asp:Panel>
    <asp:SqlDataSource ID="sdsConsSheetB_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT b.SheetNo, b.Items, b.SheetNoItems, b.ConsNo, c.ConsName, b.Price, b.Quantity FROM ConsSheetB AS b LEFT OUTER JOIN Consumables AS c ON c.ConsNo = b.ConsNo WHERE (b.SheetNo = @SheetNo)
order by SheetNoItems">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridConsSheetA_List" Name="SheetNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:FormView ID="fvConsSheetB_Detail" runat="server" DataKeyNames="SheetNoItems" DataSourceID="sdsConsSheetB_Detail" OnDataBound="fvConsSheetB_Detail_DataBound" Width="100%">
        <EditItemTemplate>
            <asp:Button ID="bbOK_B_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_B_Edit_Click" Text="確定" Width="120px" />
            <asp:Button ID="bbCancel_B_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
            <table class="TableSetting">
                <tr>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbItems_Edit" runat="server" CssClass="text-Right-Blue" Text="明細項次" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                        <asp:Label ID="eSheetNo_B_Edit" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        <asp:Label ID="eSheetNoItems_Edit" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbConsMane_Edit" runat="server" CssClass="text-Right-Blue" Text="申請耗材" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                        <asp:TextBox ID="eConsNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="30%" />
                        <asp:TextBox ID="eConsName_Edit" runat="server" CssClass="text-Left-Black" Enabled="false" AutoPostBack="true" OnTextChanged="eConsName_Edit_TextChanged" Text='<%# Eval("ConsName") %>' Width="65%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbItemsStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="處理狀態" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eItemStatus_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                        <asp:Label ID="eItemStatus_Edit" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbStockQty_Edit" runat="server" CssClass="text-Right-Blue" Text="庫存量" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eStockQty_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StockQty") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbPrice_Edit" runat="server" CssClass="text-Right-Blue" Text="單價" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="ePrice_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Price") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbQuantity_Edit" runat="server" CssClass="text-Right-Blue" Text="申請數量" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:TextBox ID="eQuantity_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eQuantity_Edit_TextChanged" Text='<%# Eval("Quantity") %>' Width="65%" />
                        <asp:Label ID="eConsUnit_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="25%" />
                        <asp:Label ID="eQtyMode_Edit" runat="server" Text='<%# Eval("QtyMode") %>' Visible="false" />
                        <asp:Label ID="eConsUnit_Edit" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbAmount_B_Edit" runat="server" CssClass="text-Right-Blue" Text="小計" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eAmount_B_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColBorder ColWidth-8Col MultiLine_High">
                        <asp:Label ID="lbRemarkB_Edit" runat="server" CssClass="text-Right-Blue" Text="用途說明" Width="95%" />
                    </td>
                    <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                        <asp:TextBox ID="eRemarkB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder  ColWidth-8Col">
                        <asp:Label ID="lbBuMan_B_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eBuManName_B_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                        <asp:Label ID="eBuMan_B_Edit" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbBuDate_B_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eBuDate_B_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbModifyMan_B_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eModifyManName_B_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                        <asp:Label ID="eModifyMan_B_Edit" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbModifyDate_B_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eModifyDate_B_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
            <asp:Button ID="bbOK_B_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_B_INS_Click" Text="確定" Width="120px" />
            <asp:Button ID="bbCancel_B_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
            <table class="TableSetting">
                <tr>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbItems_INS" runat="server" CssClass="text-Right-Blue" Text="明細項次" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eItems_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                        <asp:Label ID="eSheetNo_B_INS" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        <asp:Label ID="eSheetNoItems_INS" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbConsMane_INS" runat="server" CssClass="text-Right-Blue" Text="申請耗材" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                        <asp:TextBox ID="eConsNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="30%" />
                        <asp:TextBox ID="eConsName_INS" runat="server" CssClass="text-Left-Black" Enabled="false" AutoPostBack="true" OnTextChanged="eConsName_INS_TextChanged" Text='<%# Eval("ConsName") %>' Width="65%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbItemsStatus_INS" runat="server" CssClass="text-Right-Blue" Text="處理狀態" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eItemStatus_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                        <asp:Label ID="eItemStatus_INS" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbStockQty_INS" runat="server" CssClass="text-Right-Blue" Text="庫存量" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eStockQty_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StockQty") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbPrice_INS" runat="server" CssClass="text-Right-Blue" Text="單價" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="ePrice_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Price") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbQuantity_INS" runat="server" CssClass="text-Right-Blue" Text="申請數量" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:TextBox ID="eQuantity_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eQuantity_INS_TextChanged" Text='<%# Eval("Quantity") %>' Width="65%" />
                        <asp:Label ID="eConsUnit_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="25%" />
                        <asp:Label ID="eQtyMode_INS" runat="server" Text='<%# Eval("QtyMode") %>' Visible="false" />
                        <asp:Label ID="eConsUnit_INS" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbAmount_B_INS" runat="server" CssClass="text-Right-Blue" Text="小計" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eAmount_B_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColBorder ColWidth-8Col MultiLine_High">
                        <asp:Label ID="lbRemarkB_INS" runat="server" CssClass="text-Right-Blue" Text="用途說明" Width="95%" />
                    </td>
                    <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                        <asp:TextBox ID="eRemarkB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder  ColWidth-8Col">
                        <asp:Label ID="lbBuMan_B_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eBuManName_B_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                        <asp:Label ID="eBuMan_B_INS" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbBuDate_B_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eBuDate_B_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbModifyMan_B_INS" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eModifyManName_B_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                        <asp:Label ID="eModifyMan_B_INS" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbModifyDate_B_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eModifyDate_B_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
            <asp:Button ID="bbNew_B_Empty" runat="server" CssClass="button-Blue" Text="新增明細" CommandName="New" Width="120px" />
        </EmptyDataTemplate>
        <ItemTemplate>
            <asp:Button ID="bbNew_B_List" runat="server" CssClass="button-Black" Text="新增明細" CommandName="New" Width="120px" />
            <asp:Button ID="bbEdit_B_List" runat="server" CssClass="button-Blue" Text="編輯明細" CommandName="Edit" Width="120px" />
            <asp:Button ID="bbAbort_B_List" runat="server" CssClass="button-Red" Text="明細作廢" OnClick="bbAbort_B_List_Click" Width="120px" />
            <asp:Button ID="bbDelete_B_List" runat="server" CssClass="button-Red" Text="刪除明細" OnClick="bbDelete_B_List_Click" Width="120px" />
            <table class="TableSetting">
                <tr>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbItems_List" runat="server" CssClass="text-Right-Blue" Text="明細項次" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                        <asp:Label ID="eSheetNo_B_List" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        <asp:Label ID="eSheetNoItems_List" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbConsMane_List" runat="server" CssClass="text-Right-Blue" Text="申請耗材" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                        <asp:Label ID="eConsNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsNo") %>' Width="30%" />
                        <asp:Label ID="eConsName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsName") %>' Width="65%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbItemsStatus_List" runat="server" CssClass="text-Right-Blue" Text="處理狀態" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eItemStatus_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                        <asp:Label ID="eItemStatus_List" runat="server" Text='<%# Eval("ItemStatus") %>' Visible="false" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbStockQty_List" runat="server" CssClass="text-Right-Blue" Text="庫存量" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eStockQty_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StockQty") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbPrice_List" runat="server" CssClass="text-Right-Blue" Text="單價" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="ePrice_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Price") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbQuantity_List" runat="server" CssClass="text-Right-Blue" Text="申請數量" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eQuantity_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Quantity") %>' Width="65%" />
                        <asp:Label ID="eConsUnit_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ConsUnit_C") %>' Width="25%" />
                        <asp:Label ID="eQtyMode_List" runat="server" Text='<%# Eval("QtyMode") %>' Visible="false" />
                        <asp:Label ID="eConsUnit_List" runat="server" Text='<%# Eval("ConsUnit") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbAmount_B_List" runat="server" CssClass="text-Right-Blue" Text="小計" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eAmount_B_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColBorder ColWidth-8Col MultiLine_High">
                        <asp:Label ID="lbRemarkB_List" runat="server" CssClass="text-Right-Blue" Text="用途說明" Width="95%" />
                    </td>
                    <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                        <asp:TextBox ID="eRemarkB_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder  ColWidth-8Col">
                        <asp:Label ID="lbBuMan_B_List" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eBuManName_B_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                        <asp:Label ID="eBuMan_B_List" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbBuDate_B_List" runat="server" CssClass="text-Right-Blue" Text="建檔日" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eBuDate_B_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbModifyMan_B_List" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eModifyManName_B_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                        <asp:Label ID="eModifyMan_B_List" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbModifyDate_B_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eModifyDate_B_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
    <asp:SqlDataSource ID="sdsConsSheetB_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT b.SheetNo, b.Items, b.SheetNoItems, b.ConsNo, c.ConsName, b.Price, b.Quantity, c.StockQty, b.Amount, b.RemarkB, b.QtyMode, 
       b.ItemStatus, d.CLASSTXT AS ItemStatus_C, b.BuMan, e1.NAME AS BuManName, b.BuDate, b.ModifyMan, e2.NAME AS ModifyManName, b.ModifyDate,
       b.ConsUnit, d2.ClassTxt as ConsUnit_C
  FROM ConsSheetB AS b LEFT OUTER JOIN Consumables AS c ON c.ConsNo = b.ConsNo 
                       LEFT OUTER JOIN EMPLOYEE AS e1 ON e1.EMPNO = b.BuMan 
                       LEFT OUTER JOIN EMPLOYEE AS e2 ON e2.EMPNO = b.ModifyMan 
                       LEFT OUTER JOIN DBDICB AS d ON d.CLASSNO = b.ItemStatus AND d.FKEY = '總務課耗材進出單ConsSheetB      ItemStatus' 
                       LEFT OUTER JOIN DBDICB AS d2 ON d2.ClassNo = b.ConsUnit AND d.FKey = '耗材庫存        CONSUMABLES     ConsUnit'
WHERE (b.SheetNoItems = @SheetNoItems)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridConsSheetB_List" Name="SheetNoItems" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
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
            <LocalReport ReportPath="Report\ConsSheet_OrderP.rdlc" OnSubreportProcessing="Unnamed_SubreportProcessing">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
