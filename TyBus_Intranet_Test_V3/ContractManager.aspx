<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ContractManager.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ContractManager" %>

<asp:Content ID="ContractManagerForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">合約書管理</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbWorkDep_Search" runat="server" CssClass="text-Right-Blue" Text="承辦單位：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:DropDownList ID="ddlWorkDep_Search" runat="server" CssClass="button-Blue" Width="95%" DataSourceID="sdsWorkDepList" DataTextField="DepName" DataValueField="DepNo"></asp:DropDownList>
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbResponsibleDep_Search" runat="server" CssClass="text-Right-Blue" Text="主責單位：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:DropDownList ID="ddlResponsibleDep_Search" runat="server" CssClass="text-Left-Blue" Width="95%" DataSourceID="sdsResponsibleDepList" DataTextField="DepName" DataValueField="DepNo"></asp:DropDownList>
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbCustomName_Search" runat="server" CssClass="text-Right-Blue" Text="對方名稱" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eCustomName_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col" colspan="3" />
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出EXCEL" OnClick="bbExcel_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridContractManager" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#999999" BorderStyle="None" BorderWidth="1px" CellPadding="3" CssClass="titleText-S-Black" DataKeyNames="IndexNo" DataSourceID="sdsContractManagerList" GridLines="Vertical" Width="100%">
            <AlternatingRowStyle BackColor="Gainsboro" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="IndexNo" HeaderText="IndexNo" ReadOnly="True" SortExpression="IndexNo" Visible="False" />
                <asp:BoundField DataField="WorkDep" HeaderText="WorkDep" SortExpression="WorkDep" Visible="False" />
                <asp:BoundField DataField="WorkDepName" HeaderText="承辦單位" ReadOnly="True" SortExpression="WorkDepName" />
                <asp:BoundField DataField="ResponsibleDep" HeaderText="ResponsibleDep" SortExpression="ResponsibleDep" Visible="False" />
                <asp:BoundField DataField="ResponsibleDepName" HeaderText="主責單位" ReadOnly="True" SortExpression="ResponsibleDepName" />
                <asp:BoundField DataField="ContractNo" HeaderText="編號" SortExpression="ContractNo" />
                <asp:BoundField DataField="InvoiceNo" HeaderText="對方統編" SortExpression="InvoiceNo" />
                <asp:BoundField DataField="CustomName" HeaderText="對方名稱" SortExpression="CustomName" />
                <asp:BoundField DataField="ContractText" HeaderText="合約內容" SortExpression="ContractText" />
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
        <asp:FormView ID="fvContractManager" runat="server" DataKeyNames="IndexNo" DataSourceID="sdsContractManagerDetail" Width="100%" OnDataBound="fvContractManager_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_Edit_Click" Text="確定" Width="120px" UseSubmitBehavior="False" OnClientClick="this.disabled=true;" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbContractNo_Edit" runat="server" CssClass="text-Right-Blue" Text="編號：" Width="95%" />
                                    <asp:Label ID="eIndexNo_Edit" runat="server" Text='<%# Eval("IndexNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eContractNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContractNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbWorkDep_Edit" runat="server" CssClass="text-Right-Blue" Text="承辦單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eWorkDep_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eWorkDep_Edit_TextChanged" Text='<%# Eval("WorkDep") %>' Width="35%" />
                                    <asp:Label ID="eWorkDepName_Edit" runat="server" CssClass="text-Left-Black" Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbResponsibleDep_Edit" runat="server" CssClass="text-Right-Blue" Text="主責單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eResponsibleDep_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eResponsibleDep_Edit_TextChanged" Text='<%# Eval("ResponsibleDep") %>' Width="35%" />
                                    <asp:Label ID="eResponsibleDepName_Edit" runat="server" CssClass="text-Left-Black" Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:CheckBox ID="cbIsClose_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbIsClose_Edit_CheckedChanged" Text="已結案" />
                                    <asp:Label ID="eIsClose_Edit" runat="server" Text='<%# Eval("IsClose") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbInvoiceNo_Edit" runat="server" CssClass="text-Right-Blue" Text="對方統一編號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eInvoiceNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InvoiceNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbCustomName_Edit" runat="server" CssClass="text-Right-Blue" Text="對方名稱：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eCustomName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CustomName") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_Low ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbContractText_Edit" runat="server" CssClass="text-Right-Blue" Text="工程地點、內容或名稱：" Width="95%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-6Col" colspan="5">
                                    <asp:TextBox ID="eContractText_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("ContractText") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbAssignDate_Edit" runat="server" CssClass="text-Right-Blue" Text="立契日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eAssignDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbEffectiveDateS_Edit" runat="server" CssClass="text-Right-Blue" Text="合約時間(起)：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eEffectiveDateS_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EffectiveDateS","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbEffectiveDateE_Edit" runat="server" CssClass="text-Right-Blue" Text="合約時間(迄)：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eEffectiveDateE_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EffectiveDateE","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="憑證金額：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eAmount_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAmount_Edit_TextChanged" Text='<%# Eval("Amount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbTax_Edit" runat="server" CssClass="text-Right-Blue" Text="憑證稅額：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eTax_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAmount_Edit_TextChanged" Text='<%# Eval("Tax") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" rowspan="2">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" rowspan="2">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="95%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbTotalAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="合計：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eTotalAmount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TotalAmount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbStampDuty_Edit" runat="server" CssClass="text-Right-Blue" Text="印花稅：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eStampDuty_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StampDuty") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                            </tr>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_INS_Click" Text="確定" Width="120px" OnClientClick="this.disabled=true;" UseSubmitBehavior="False" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbContractNo_INS" runat="server" CssClass="text-Right-Blue" Text="編號：" Width="95%" />
                                    <asp:Label ID="eIndexNo_INS" runat="server" Text='<%# Eval("IndexNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eContractNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContractNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbWorkDep_INS" runat="server" CssClass="text-Right-Blue" Text="承辦單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eWorkDep_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eWorkDep_INS_TextChanged" Text='<%# Eval("WorkDep") %>' Width="35%" />
                                    <asp:Label ID="eWorkDepName_INS" runat="server" CssClass="text-Left-Black" Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbResponsibleDep_INS" runat="server" CssClass="text-Right-Blue" Text="主責單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eResponsibleDep_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eResponsibleDep_INS_TextChanged" Text='<%# Eval("ResponsibleDep") %>' Width="35%" />
                                    <asp:Label ID="eResponsibleDepName_INS" runat="server" CssClass="text-Left-Black" Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:CheckBox ID="cbIsClose_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbIsClose_INS_CheckedChanged" Text="已結案" />
                                    <asp:Label ID="eIsClose_INS" runat="server" Text='<%# Eval("IsClose") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbInvoiceNo_INS" runat="server" CssClass="text-Right-Blue" Text="對方統一編號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eInvoiceNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InvoiceNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbCustomName_INS" runat="server" CssClass="text-Right-Blue" Text="對方名稱：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eCustomName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CustomName") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_Low ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbContractText_INS" runat="server" CssClass="text-Right-Blue" Text="工程地點、內容或名稱：" Width="95%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-6Col" colspan="5">
                                    <asp:TextBox ID="eContractText_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("ContractText") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbAssignDate_INS" runat="server" CssClass="text-Right-Blue" Text="立契日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eAssignDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbEffectiveDateS_INS" runat="server" CssClass="text-Right-Blue" Text="合約時間(起)：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eEffectiveDateS_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EffectiveDateS","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbEffectiveDateE_INS" runat="server" CssClass="text-Right-Blue" Text="合約時間(迄)：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eEffectiveDateE_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EffectiveDateE","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbAmount_INS" runat="server" CssClass="text-Right-Blue" Text="憑證金額：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eAmount_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAmount_INS_TextChanged" Text='<%# Eval("Amount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbTax_INS" runat="server" CssClass="text-Right-Blue" Text="憑證稅額：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eTax_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAmount_INS_TextChanged" Text='<%# Eval("Tax") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" rowspan="2">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" rowspan="2">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="95%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbTotalAmount_INS" runat="server" CssClass="text-Right-Blue" Text="合計：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eTotalAmount_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TotalAmount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbStampDuty_INS" runat="server" CssClass="text-Right-Blue" Text="印花稅：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eStampDuty_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StampDuty") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                            </tr>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CausesValidation="false" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CssClass="button-Blue" CommandName="Edit" Text="修改" Width="120px" />                
                &nbsp;<asp:Button ID="bbDelete_List" runat="server" CausesValidation="False" CssClass="button-Red" OnClick="bbDelete_List_Click" Text="刪除" Width="120px" Visible="false" />
                &nbsp;<asp:Button ID="bbCaseClose_List" runat="server" CausesValidation="false" CssClass="button-Red" OnClick="bbCaseClose_List_Click" Text="結案" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbContractNo_List" runat="server" CssClass="text-Right-Blue" Text="編號：" Width="95%" />
                                    <asp:Label ID="eIndexNo_List" runat="server" Text='<%# Eval("IndexNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eContractNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContractNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbWorkDep_List" runat="server" CssClass="text-Right-Blue" Text="承辦單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eWorkDep_List" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eWorkDep_List_TextChanged" Text='<%# Eval("WorkDep") %>' Width="35%" />
                                    <asp:Label ID="eWorkDepName_List" runat="server" CssClass="text-Left-Black" Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbResponsibleDep_List" runat="server" CssClass="text-Right-Blue" Text="主責單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eResponsibleDep_List" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eResponsibleDep_List_TextChanged" Text='<%# Eval("ResponsibleDep") %>' Width="35%" />
                                    <asp:Label ID="eResponsibleDepName_List" runat="server" CssClass="text-Left-Black" Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:CheckBox ID="cbIsClose_List" runat="server" CssClass="text-Left-Black" Text="已結案" />
                                    <asp:Label ID="eIsClose_List" runat="server" Text='<%# Eval("IsClose") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbInvoiceNo_List" runat="server" CssClass="text-Right-Blue" Text="對方統一編號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eInvoiceNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InvoiceNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbCustomName_List" runat="server" CssClass="text-Right-Blue" Text="對方名稱：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eCustomName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CustomName") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_Low ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbContractText_List" runat="server" CssClass="text-Right-Blue" Text="工程地點、內容或名稱：" Width="95%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-6Col" colspan="5">
                                    <asp:TextBox ID="eContractText_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("ContractText") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbAssignDate_List" runat="server" CssClass="text-Right-Blue" Text="立契日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eAssignDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbEffectiveDateS_List" runat="server" CssClass="text-Right-Blue" Text="合約時間(起)：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eEffectiveDateS_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EffectiveDateS","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbEffectiveDateE_List" runat="server" CssClass="text-Right-Blue" Text="合約時間(迄)：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eEffectiveDateE_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EffectiveDateE","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbAmount_List" runat="server" CssClass="text-Right-Blue" Text="憑證金額：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Amount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbTax_List" runat="server" CssClass="text-Right-Blue" Text="憑證稅額：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eTax_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Tax") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" rowspan="2">
                                    <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" rowspan="2">
                                    <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Width="95%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbTotalAmount_List" runat="server" CssClass="text-Right-Blue" Text="合計：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eTotalAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TotalAmount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbStampDuty_List" runat="server" CssClass="text-Right-Blue" Text="印花稅：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eStampDuty_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StampDuty") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                                <td class="ColWidth-6Col" />
                            </tr>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsContractManagerList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT IndexNo, ContractNo, WorkDep, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = cm.WorkDep)) AS WorkDepName, ResponsibleDep, (SELECT NAME FROM DEPARTMENT AS DEPARTMENT_1 WHERE (DEPNO = cm.ResponsibleDep)) AS ResponsibleDepName, InvoiceNo, CustomName, ContractText FROM ContractManager AS cm WHERE 1 &lt;&gt; 1"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsContractManagerDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" DeleteCommand="DELETE FROM ContractManager WHERE (IndexNo = @IndexNo)" InsertCommand="INSERT INTO ContractManager(IndexNo, ContractNo, WorkDep, ResponsibleDep, InvoiceNo, CustomName, ContractText, AssignDate, EffectiveDateS, EffectiveDateE, IsClose, Amount, Tax, TotalAmount, StampDuty, Remark, BuMan, BuDate) VALUES (@IndexNo, @ContractNo, @WorkDep, @ResponsibleDep, @InvoiceNo, @CustomName, @ContractText, @AssignDate, @EffectiveDateS, @EffectiveDateE, @IsClose, @Amount, @Tax, @TotalAmount, @StampDuty, @Remark, @BuMan, @BuDate)" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT IndexNo, ContractNo, WorkDep, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = cm.WorkDep)) AS WorkDepName, ResponsibleDep, (SELECT NAME FROM DEPARTMENT AS DEPARTMENT_1 WHERE (DEPNO = cm.ResponsibleDep)) AS ResponsibleDepName, InvoiceNo, CustomName, ContractText, AssignDate, EffectiveDateS, EffectiveDateE, IsClose, Amount, Tax, TotalAmount, StampDuty, Remark, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = cm.BuMan)) AS BuManName, BuDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = cm.ModifyMan)) AS ModifyManName, ModifyDate FROM ContractManager AS cm WHERE (IndexNo = @IndexNo)" UpdateCommand="UPDATE ContractManager SET ContractNo = @ContractNo, WorkDep = @WorkDep, ResponsibleDep = @ResponsibleDep, InvoiceNo = @InvoiceNo, CustomName = @CustomName, ContractText = @ContractText, AssignDate = @AssignDate, EffectiveDateS = @EffectiveDateS, EffectiveDateE = @EffectiveDateE, IsClose = @IsClose, Amount = @Amount, Tax = @Tax, TotalAmount = @TotalAmount, StampDuty = @StampDuty, Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate WHERE (IndexNo = @IndexNo)">
        <DeleteParameters>
            <asp:Parameter Name="IndexNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="IndexNo" />
            <asp:Parameter Name="ContractNo" />
            <asp:Parameter Name="WorkDep" />
            <asp:Parameter Name="ResponsibleDep" />
            <asp:Parameter Name="InvoiceNo" />
            <asp:Parameter Name="CustomName" />
            <asp:Parameter Name="ContractText" />
            <asp:Parameter Name="AssignDate" />
            <asp:Parameter Name="EffectiveDateS" />
            <asp:Parameter Name="EffectiveDateE" />
            <asp:Parameter Name="IsClose" />
            <asp:Parameter Name="Amount" />
            <asp:Parameter Name="Tax" />
            <asp:Parameter Name="TotalAmount" />
            <asp:Parameter Name="StampDuty" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="BuDate" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridContractManager" Name="IndexNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="ContractNo" />
            <asp:Parameter Name="WorkDep" />
            <asp:Parameter Name="ResponsibleDep" />
            <asp:Parameter Name="InvoiceNo" />
            <asp:Parameter Name="CustomName" />
            <asp:Parameter Name="ContractText" />
            <asp:Parameter Name="AssignDate" />
            <asp:Parameter Name="EffectiveDateS" />
            <asp:Parameter Name="EffectiveDateE" />
            <asp:Parameter Name="IsClose" />
            <asp:Parameter Name="Amount" />
            <asp:Parameter Name="Tax" />
            <asp:Parameter Name="TotalAmount" />
            <asp:Parameter Name="StampDuty" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="IndexNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsWorkDepList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="select cast(null as varchar) DepNo, cast(null as varchar) DepName"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsResponsibleDepList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="select cast(null as varchar) DepNo, cast(null as varchar) DepName"></asp:SqlDataSource>
</asp:Content>
