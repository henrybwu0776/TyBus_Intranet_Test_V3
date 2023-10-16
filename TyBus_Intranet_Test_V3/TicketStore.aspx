<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="TicketStore.aspx.cs" Inherits="TyBus_Intranet_Test_V3.TicketStore" %>

<asp:Content ID="TicketStoreForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">票證庫存庫位維護</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="歸屬單位" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_TextChanged" Width="40%" />
                    <asp:Label ID="lbDepName" runat="server" CssClass="text-Left-Black" Width="50%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbAdminEmpNo" runat="server" CssClass="text-Right-Blue" Text="管理人" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eAdminEmpNo" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAdminEmpNo_TextChanged" Width="40%" />
                    <asp:Label ID="eAdminEmpName" runat="server" CssClass="text-Left-Black" Width="50%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:RadioButtonList ID="rbSearchMode" runat="server" CssClass="text-Left-Black" AutoPostBack="true" RepeatColumns="3" Width="95%">
                        <asp:ListItem Text="使用中庫位" Value="V" Selected="True" />
                        <asp:ListItem Text="已停用庫位" Value="X" />
                        <asp:ListItem Text="全部庫位" Value="Z" />
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="95%" />
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
    <asp:Panel ID="plDataList" runat="server" CssClass="ShowPanel">
        <asp:FormView ID="fvTicketStoreDetail" runat="server" DataKeyNames="WarehouseNo" DataSourceID="sdsTicketStoreDetail" Width="100%" OnDataBound="fvTicketStoreDetail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Update" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Update" Text="確定" />
                &nbsp;<asp:Button ID="bbCancel_Update" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbWarehouseNo_Edit" runat="server" CssClass="text-Right-Blue" Text="庫位編號" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eWarehouseNo_Edit" runat="server" Text='<%# Eval("WarehouseNo") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="所屬單位" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_Edit" runat="server" AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Text='<%# Bind("DepNo") %>' />
                                    <asp:TextBox ID="eDepName_Edit" runat="server" Enabled="false" Text='<%# Bind("DepName") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbWarehousePosition_Edit" runat="server" CssClass="text-Right-Blue" Text="所在位置" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eWarehousePosition_Edit" runat="server" Text='<%# Bind("WarehousePosition") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAdminEmpNo_Edit" runat="server" CssClass="text-Right-Blue" Text="保管人" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eAdminEmpNo_Edit" runat="server" AutoPostBack="true" OnTextChanged="eAdminEmpNo_Edit_TextChanged" Text='<%# Bind("AdminEmpNo") %>' />
                                    <asp:TextBox ID="eAdminEmpName_Edit" runat="server" Enabled="false" Text='<%# Bind("AdminEmpName") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbIsInused_Edit" runat="server" CssClass="text-Right-Blue" Text="庫位狀況" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:RadioButtonList ID="rbIsInused_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnSelectedIndexChanged="rbIsInused_Edit_SelectedIndexChanged" Width="95%" RepeatColumns="2">
                                        <asp:ListItem Text="使用中" Value="V" />
                                        <asp:ListItem Text="已停用" Value="X" />
                                    </asp:RadioButtonList>
                                    <asp:TextBox ID="eIsInused_Edit" runat="server" Text='<%# Bind("IsInused") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" rowspan="3">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2" rowspan="3">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" TextMode="MultiLine" Width="95%" Height="95%" Text='<%# Bind("Remark") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_Edit" runat="server" Text='<%# Eval("BuMan") %>' />
                                    <asp:Label ID="eBuMan_C_Edit" runat="server" Text='<%# Eval("BuMan_C") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="BuDateLabel" runat="server" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" Text='<%# Eval("ModifyMan") %>' />
                                    <asp:Label ID="eModifyMan_C_Edit" runat="server" Text='<%# Eval("ModifyMan_C") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' />
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
            <EmptyDataTemplate>
                <asp:Button ID="bbInsert_Empty" runat="server" CommandName="new" CssClass="button-Black" Text="新增庫位" />
            </EmptyDataTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_Insert" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Insert" Text="確定" />
                &nbsp;<asp:Button ID="bbCancel_Insert" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbWarehouseNo_Insert" runat="server" CssClass="text-Right-Blue" Text="庫位編號" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eWarehouseNo_Insert" runat="server" Text='<%# Bind("WarehouseNo") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Insert" runat="server" CssClass="text-Right-Blue" Text="所屬單位" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_Insert" runat="server" AutoPostBack="true" OnTextChanged="eDepNo_Insert_TextChanged" Text='<%# Bind("DepNo") %>' />
                                    <asp:TextBox ID="eDepName_Insert" runat="server" Enabled="false" Text='<%# Bind("DepName") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbWarehousePosition_Insert" runat="server" CssClass="text-Right-Blue" Text="所在位置" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eWarehousePosition_Insert" runat="server" Text='<%# Bind("WarehousePosition") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAdminEmpNo_Insert" runat="server" CssClass="text-Right-Blue" Text="保管人" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eAdminEmpNo_Insert" runat="server" AutoPostBack="true" OnTextChanged="eAdminEmpNo_Insert_TextChanged" Text='<%# Bind("AdminEmpNo") %>' />
                                    <asp:TextBox ID="eAdminEmpName_Insert" runat="server" Enabled="false" Text='<%# Bind("AdminEmpName") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbIsInused_Insert" runat="server" CssClass="text-Right-Blue" Text="庫位狀況" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:RadioButtonList ID="rbIsInused_Insert" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnSelectedIndexChanged="rbIsInused_Insert_SelectedIndexChanged" Width="95%" RepeatColumns="2">
                                        <asp:ListItem Text="使用中" Value="V" />
                                        <asp:ListItem Text="已停用" Value="X" />
                                    </asp:RadioButtonList>
                                    <asp:TextBox ID="eIsInused_Insert" runat="server" Text='<%# Bind("IsInused") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" rowspan="3">
                                    <asp:Label ID="lbRemark_Insert" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2" rowspan="3">
                                    <asp:TextBox ID="eRemark_Insert" runat="server" TextMode="MultiLine" Width="95%" Height="95%" Text='<%# Bind("Remark") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_Insert" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_Insert" runat="server" Text='<%# Eval("BuMan") %>' />
                                    <asp:Label ID="eBuMan_C_Insert" runat="server" Text='<%# Eval("BuMan_C") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_Insert" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="BuDateLabel" runat="server" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_Insert" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eModifyMan_Insert" runat="server" Text='<%# Eval("ModifyMan") %>' />
                                    <asp:Label ID="eModifyMan_C_Insert" runat="server" Text='<%# Eval("ModifyMan_C") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_Insert" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_Insert" runat="server" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' />
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
            <ItemTemplate>
                <asp:Button ID="bbInsert" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增庫位" />
                &nbsp;<asp:Button ID="bbEdit" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="Edit" Text="編輯庫位" />
                &nbsp;<asp:Button ID="bbDelete" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Delete" Text="刪除庫位" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbWarehouseNo_List" runat="server" CssClass="text-Right-Blue" Text="庫位編號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eWarehouseNo_List" runat="server" Text='<%# Eval("WarehouseNo") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="所屬單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" Text='<%# Eval("DepNo") %>' />
                            <asp:Label ID="eDepName_List" runat="server" Text='<%# Eval("DepName") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbWarehousePosition_List" runat="server" CssClass="text-Right-Blue" Text="所在位置" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eWarehousePosition_List" runat="server" Text='<%# Eval("WarehousePosition") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAdminEmpNo_List" runat="server" CssClass="text-Right-Blue" Text="保管人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eAdminEmpNo_List" runat="server" Text='<%# Eval("AdminEmpNo") %>' />
                            <asp:Label ID="eAdminEmpName_List" runat="server" Text='<%# Eval("AdminEmpName") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbIsInused_List" runat="server" CssClass="text-Right-Blue" Text="庫位狀況" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eIsInused_List" runat="server" Text='<%# Eval("IsInused_C") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="3">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2" rowspan="3">
                            <asp:Label ID="eRemark_List" runat="server" Width="95%" Height="95%" Text='<%# Eval("Remark") %>' />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuMan_List" runat="server" Text='<%# Eval("BuMan") %>' />
                            <asp:Label ID="eBuMan_C_List" runat="server" Text='<%# Eval("BuMan_C") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="BuDateLabel" runat="server" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eModifyMan_List" runat="server" Text='<%# Eval("ModifyMan") %>' />
                            <asp:Label ID="eModifyMan_C_List" runat="server" Text='<%# Eval("ModifyMan_C") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_List" runat="server" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' />
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
        <asp:GridView ID="gridTicketStoreList" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="WarehouseNo" DataSourceID="sdsTicketStoreList"
            ForeColor="#333333" GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="WarehouseNo" HeaderText="庫位編號" ReadOnly="True" SortExpression="WarehouseNo" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="歸屬單位" SortExpression="DepName" />
                <asp:BoundField DataField="AdminEmpNo" HeaderText="AdminEmpNo" SortExpression="AdminEmpNo" Visible="False" />
                <asp:BoundField DataField="AdminEmpName" HeaderText="庫位管理人" SortExpression="AdminEmpName" />
                <asp:BoundField DataField="WarehousePosition" HeaderText="所在位置" SortExpression="WarehousePosition" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" />
                <asp:BoundField DataField="BuDate" DataFormatString="{0:d}" HeaderText="建檔日期" SortExpression="BuDate" />
                <asp:BoundField DataField="BuMan" HeaderText="BuMan" SortExpression="BuMan" Visible="False" />
                <asp:BoundField DataField="BuMan_C" HeaderText="建檔人" ReadOnly="True" SortExpression="BuMan_C" />
                <asp:BoundField DataField="ModifyDate" DataFormatString="{0:d}" HeaderText="異動日期" SortExpression="ModifyDate" />
                <asp:BoundField DataField="ModifyMan" HeaderText="ModifyMan" SortExpression="ModifyMan" Visible="False" />
                <asp:BoundField DataField="ModifyMan_C" HeaderText="異動人" ReadOnly="True" SortExpression="ModifyMan_C" />
                <asp:BoundField DataField="IsInused" HeaderText="使用情況" SortExpression="IsInused" />
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
    </asp:Panel>
    <asp:SqlDataSource ID="sdsTicketStoreList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT WarehouseNo, DepNo, DepName, AdminEmpNo, AdminEmpName, WarehousePosition, Remark, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuMan_C, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyMan_C, IsInused FROM TicketStore AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsTicketStoreDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        DeleteCommand="DELETE FROM TicketStore WHERE (WarehouseNo = @WarehouseNo)"
        InsertCommand="INSERT INTO TicketStore(WarehouseNo, DepNo, DepName, AdminEmpNo, AdminEmpName, WarehousePosition, Remark, BuDate, BuMan, IsInused) VALUES (@WarehouseNo, @DepNo, @DepName, @AdminEmpNo, @AdminEmpName, @WarehousePosition, @Remark, @BuDate, @BuMan, @IsInused)"
        SelectCommand="SELECT WarehouseNo, DepNo, DepName, AdminEmpNo, AdminEmpName, WarehousePosition, Remark, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuMan_C, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyMan_C, IsInused, CASE WHEN isnull(IsInused , 'X') = 'V' THEN '使用中' ELSE '己停用' END AS IsInused_C FROM TicketStore AS a WHERE (WarehouseNo = @WarehouseNo)"
        UpdateCommand="UPDATE TicketStore SET DepNo = @DepNo, DepName = @DepName, AdminEmpNo = @AdminEmpNo, AdminEmpName = @AdminEmpName, WarehousePosition = @WarehousePosition, Remark = @Remark, ModifyDate = @ModifyDate, ModifyMan = @ModifyMan, IsInused = @IsInused WHERE (WarehouseNo = @WarehouseNo)"
        OnDeleted="sdsTicketStoreDetail_Deleted"
        OnInserted="sdsTicketStoreDetail_Inserted"
        OnInserting="sdsTicketStoreDetail_Inserting"
        OnUpdated="sdsTicketStoreDetail_Updated"
        OnUpdating="sdsTicketStoreDetail_Updating">
        <DeleteParameters>
            <asp:Parameter Name="WarehouseNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="WarehouseNo" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="DepName" />
            <asp:Parameter Name="AdminEmpNo" />
            <asp:Parameter Name="AdminEmpName" />
            <asp:Parameter Name="WarehousePosition" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuDate" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="IsInused" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridTicketStoreList" Name="WarehouseNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="DepName" />
            <asp:Parameter Name="AdminEmpNo" />
            <asp:Parameter Name="AdminEmpName" />
            <asp:Parameter Name="WarehousePosition" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="WarehouseNo" />
            <asp:Parameter Name="IsInused" />
        </UpdateParameters>
    </asp:SqlDataSource>
</asp:Content>
