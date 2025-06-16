<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ConsSheet_Fix.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ConsSheet_Fix" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="ConsSheet_FixForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">總務請修單</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="請修單位" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eDepName_Search" runat="server" CssClass="text-Left-Black" Width="60%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbAssignMan_Search" runat="server" CssClass="text-Right-Blue" Text="申請人" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                    <asp:TextBox ID="eAssignMan_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignMan_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eAssignManName_Search" runat="server" CssClass="text-Left-Black" Width="60%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbBuDate_Search" runat="server" CssClass="text-Right-Blue" Text="申請日期" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                    <asp:TextBox ID="eBuDateS_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="eSplit_1" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eBuDateE_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-10Col" colspan="6">
                    <asp:Label ID="eErrorMSG_Main" runat="server" CssClass="errorMessageText" Text="" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Blue" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" OnClick="bbExit_Click" Width="95%" />
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
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <div>
        <asp:Label ID="eErrorMSG_A" runat="server" CssClass="errorMessageText" Text="" Visible="false" />
    </div>
    <asp:Panel ID="plShowData_A" runat="server" CssClass="ShowPanel" >
        <asp:GridView ID="gridSheetA" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="SheetNo" DataSourceID="sdsSheetA_List" GridLines="None" PageSize="5" Width="100%">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="BuDate" HeaderText="申請日期" SortExpression="BuDate" />
                <asp:BoundField DataField="SheetNo" HeaderText="請修單號" ReadOnly="True" SortExpression="SheetNo" />
                <asp:BoundField DataField="DepNo" HeaderText="單位代碼" SortExpression="DepNo" />
                <asp:BoundField DataField="DepName" HeaderText="申請單位" SortExpression="DepName" />
                <asp:BoundField DataField="AssignMan" HeaderText="申請人工號" SortExpression="AssignMan" />
                <asp:BoundField DataField="AssignManName" HeaderText="申請人姓名" SortExpression="AssignManName" />
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
        <asp:FormView ID="fvSheetA_Detail" runat="server" DataKeyNames="SheetNo" DataSourceID="sdsSheetA_Detail" OnDataBound="fvSheetA_Detail_DataBound" Width="100%">
            <EditItemTemplate>
                <asp:Button ID="bbOKA_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKA_Edit_Click" Text="確定" Width="120px" />
                <asp:Button ID="bbCancelA_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetNoA_Edit" runat="server" CssClass="text-Right-Blue" Text="請修單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetNoA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateA_Edit" runat="server" CssClass="text-Right-Blue" Text="申請報修日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eBuDateA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAssignMan_Edit" runat="server" CssClass="text-Right-Blue" Text="申請人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eAssignMan_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignMan_Edit_TextChanged" Text='<%# Eval("AssignMan") %>' Width="35%" />
                            <asp:Label ID="eAssignManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lnBuManA_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuManA_Edit" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                            <asp:Label ID="eBuManNameA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3">
                            <asp:Label ID="lbRemarkA_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3" colspan="7">
                            <asp:TextBox ID="eRemarkA_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManA_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyManA_Edit" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                            <asp:Label ID="eModifyManNameA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateA_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateA_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
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
                            <asp:Label ID="lbSheetNoA_INS" runat="server" CssClass="text-Right-Blue" Text="請修單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetNoA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateA_INS" runat="server" CssClass="text-Right-Blue" Text="申請報修日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eBuDateA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="申請單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAssignMan_INS" runat="server" CssClass="text-Right-Blue" Text="申請人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eAssignMan_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignMan_INS_TextChanged" Text='<%# Eval("AssignMan") %>' Width="35%" />
                            <asp:Label ID="eAssignManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lnBuManA_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuManA_INS" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                            <asp:Label ID="eBuManNameA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3">
                            <asp:Label ID="lbRemarkA_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3" colspan="7">
                            <asp:TextBox ID="eRemarkA_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkA") %>' Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManA_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyManA_INS" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                            <asp:Label ID="eModifyManNameA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateA_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateA_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
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
                <asp:Button ID="bbPrint_List"  runat="server" CssClass="button-Black" OnClick="bbPrint_List_Click" Text="列印請修單" Width="120px" />
                <asp:Button ID="bbAbortA_List" runat="server" CssClass="button-Blue" OnClick="bbAbortA_List_Click" Text="請修單作廢" Width="120px" />
                <asp:Button ID="bbDelete_List" runat="server" CssClass="button-Red" OnClick="bbDelete_List_Click" Text="刪除" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbSheetNoA_List" runat="server" CssClass="text-Right-Blue" Text="請修單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eSheetNoA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SheetNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuDateA_List" runat="server" CssClass="text-Right-Blue" Text="申請報修日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuDateA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate") %>' Width="95%" />
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
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lnBuManA_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuManA_List" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                            <asp:Label ID="eBuManNameA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3">
                            <asp:Label ID="lbRemarkA_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" rowspan="3" colspan="7">
                            <asp:TextBox ID="eRemarkA_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("RemarkA") %>' Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyManA_List" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyManA_List" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                            <asp:Label ID="eModifyManNameA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDateA_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyDateA_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
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
    <asp:SqlDataSource ID="sdsSheetA_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select a.SheetNo, a.DepNo, d.[Name] DepName, a.AssignMan, e.[Name] AssignManName, convert(varchar(10), a.BuDate, 111) BuDate 
  from ConsSheetA a left join Department d on d.DepNo = a.DepNo 
                    left join Employee e on e.EmpNo = a.AssignMan 
 where isnull(a.SheetMode, '') = ''"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsSheetA_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select a.SheetNo, a.DepNo, d.[Name] DepName, a.AssignMan, e1.[Name] AssignManName, convert(varchar(10), a.BuDate, 111) BuDate, a.RemarkA, a.ModifyMan, e2.[Name] ModifyManName, convert(varchar(10), a.ModifyDate, 111) ModifyDate, a.BuMan, e3.[Name] BuManName
from ConsSheetA a left join Department d on d.DepNo = a.DepNo
left join Employee e1 on e1.EmpNo = a.AssignMan
left join Employee e2 on e2.EmpNo = a.ModifyMan
left join Employee e3 on e3.EmpNo = a.BuMan
where a.SheetNo = @SheetNo">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridSheetA" Name="SheetNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <div>
        <asp:Label ID="eErrorMSG_B" runat="server" CssClass="errorMessageText" Visible="false" />
    </div>
    <asp:Panel ID="plShowData_B" runat="server" CssClass="ShowPanel-Detail_C">
        <asp:GridView ID="gridSheetB_List" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#DEDFDE" BorderStyle="None" BorderWidth="1px" CellPadding="4" DataKeyNames="SheetNoItems" DataSourceID="sdsSheetB_List" ForeColor="Black" GridLines="Vertical" PageSize="5">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="SheetNoItems" HeaderText="SheetNoItems" ReadOnly="True" SortExpression="SheetNoItems" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                <asp:BoundField DataField="RemarkB" HeaderText="請修內容" SortExpression="RemarkB" />
                <asp:BoundField DataField="ItemStatus_C" HeaderText="執行狀態" SortExpression="ItemStatus_C" />
            </Columns>
            <FooterStyle BackColor="#CCCC99" />
            <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
            <RowStyle BackColor="#F7F7DE" />
            <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#FBFBF2" />
            <SortedAscendingHeaderStyle BackColor="#848384" />
            <SortedDescendingCellStyle BackColor="#EAEAD3" />
            <SortedDescendingHeaderStyle BackColor="#575357" />
        </asp:GridView>
        </asp:Panel>
        <asp:Panel ID="plShowData_C" runat="server" CssClass="ShowPanel-Detail_B">
        <asp:FormView ID="fvSheetB_Detail" runat="server" Width="100%" DataKeyNames="SheetNoItems" DataSourceID="sdsSheetB_Detail" OnDataBound="fvSheetB_Detail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOKB_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKB_Edit_Click" Text="確定" Width="120px" />
                <asp:Button ID="UpdateCancelButton" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbItems_Edit" runat="server" CssClass="text-Right-Blue" Text="項次" Width="95%" />
                            <asp:Label ID="eSheetNoB_Edit" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="eItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoItems_Edit" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbItemStatus_Edit" runat="server" CssClass="text-Right-Blue" Text="狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                            <asp:Label ID="eItemStatus_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColBorder ColWidth-5Col MultiLine_High">
                            <asp:Label ID="lbRemarkB_Edit" runat="server" CssClass="text-Right-Blue" Text="報修項目" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-5Col MultiLine_High" colspan="4">
                            <asp:TextBox ID="eRemarkB_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbBuDateB_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="eBuDateB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbBuManB_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                            <asp:Label ID="eBuManB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManNameB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbModifyDateB_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="eModifyDateB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbModifyManB_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                            <asp:Label ID="eModifyManB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManNameB_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-5Col" />
                        <td class="ColWidth-5Col" />
                        <td class="ColWidth-5Col" />
                        <td class="ColWidth-5Col" />
                        <td class="ColWidth-5Col" />
                    </tr>
                </table>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOKB_INS" runat="server" CausesValidation="True" CssClass="button-Black" OnClick="bbOKB_INS_Click" Text="確定" Width="120px" />
                <asp:Button ID="bbCancelB_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbItems_INS" runat="server" CssClass="text-Right-Blue" Text="項次" Width="95%" />
                            <asp:Label ID="eSheetNoB_INS" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="eItems_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoItems_INS" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbItemStatus_INS" runat="server" CssClass="text-Right-Blue" Text="狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                            <asp:Label ID="eItemStatus_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColBorder ColWidth-5Col MultiLine_High">
                            <asp:Label ID="lbRemarkB_INS" runat="server" CssClass="text-Right-Blue" Text="報修項目" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-5Col MultiLine_High" colspan="4">
                            <asp:TextBox ID="eRemarkB_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbBuDateB_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="eBuDateB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbBuManB_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                            <asp:Label ID="eBuManB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManNameB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbModifyDateB_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="eModifyDateB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbModifyManB_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                            <asp:Label ID="eModifyManB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManNameB_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-5Col" />
                        <td class="ColWidth-5Col" />
                        <td class="ColWidth-5Col" />
                        <td class="ColWidth-5Col" />
                        <td class="ColWidth-5Col" />
                    </tr>
                </table>
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNewB_Empty" runat="server" CssClass="button-Blue" CommandName="New" Text="新增明細" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNewB_List" runat="server" CssClass="button-Black" CommandName="New" Text="新增明細" Width="120px" />
                <asp:Button ID="bbEditB_List" runat="server" CssClass="button-Blue" CommandName="Edit" Text="修改明細" Width="120px" />
                <asp:Button ID="bbAbortB_List" runat="server" CssClass="button-Black" OnClick="bbAbortB_List_Click" Text="明細作廢" Width="120px" />
                <asp:Button ID="bbDeleteB_List" runat="server" CssClass="button-Red" OnClick="bbDeleteB_List_Click" Text="刪除明細" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbItems_List" runat="server" CssClass="text-Right-Blue" Text="項次" Width="95%" />
                            <asp:Label ID="eSheetNoB_List" runat="server" Text='<%# Eval("SheetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eSheetNoItems_List" runat="server" Text='<%# Eval("SheetNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbItemStatus_List" runat="server" CssClass="text-Right-Blue" Text="狀態" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                            <asp:Label ID="eItemStatus_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ItemStatus_C") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColBorder ColWidth-5Col MultiLine_High">
                            <asp:Label ID="lbRemarkB_List" runat="server" CssClass="text-Right-Blue" Text="報修項目" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-5Col MultiLine_High" colspan="4">
                            <asp:TextBox ID="eRemarkB_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("RemarkB") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbBuDateB_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="eBuDateB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbBuManB_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                            <asp:Label ID="eBuManB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManNameB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbModifyDateB_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="eModifyDateB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col">
                            <asp:Label ID="lbModifyManB_List" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                            <asp:Label ID="eModifyManB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManNameB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-5Col" />
                        <td class="ColWidth-5Col" />
                        <td class="ColWidth-5Col" />
                        <td class="ColWidth-5Col" />
                        <td class="ColWidth-5Col" />
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
        <asp:SqlDataSource ID="sdsSheetB_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select b.SheetNoItems, b.Items, b.RemarkB, d.ClassTxt ItemStatus_C
from ConsSheetB b left join DBDICB d on d.ClassNo = b.ItemStatus and d.FKey = '總務課耗材進出單ConsSheetB      ItemStatus'
where b.SheetNo = @SheetNo">
            <SelectParameters>
                <asp:ControlParameter ControlID="gridSheetA" Name="SheetNo" PropertyName="SelectedValue" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="sdsSheetB_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select b.SheetNoItems, b.Sheetno, b.Items, b.RemarkB, d.ClassTxt ItemStatus_C, convert(varchar(10), b.BuDate, 111) BuDate, b.BuMan, e1.[Name] BuManName, convert(Varchar(10), b.ModifyDate, 111) ModifyDate, b.ModifyMan, e2.[Name] ModifyManName
from ConsSheetB b left join DBDICB d on d.ClassNo = b.ItemStatus and d.FKey = '總務課耗材進出單ConsSheetB      ItemStatus'
left join Employee e1 on e1.EmpNo = b.BuMan
left join Employee e2 on e2.EmpNo = b.ModifyMan
where SheetNoItems = @SheetNoItems">
            <SelectParameters>
                <asp:ControlParameter ControlID="gridSheetB_List" Name="SheetNoItems" PropertyName="SelectedValue" />
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
            <LocalReport ReportPath="Report\ConsSheet_FixP.rdlc" OnSubreportProcessing="Unnamed_SubreportProcessing">
            </LocalReport>
        </rsweb:ReportViewer>
        </asp:Panel>
</asp:Content>
