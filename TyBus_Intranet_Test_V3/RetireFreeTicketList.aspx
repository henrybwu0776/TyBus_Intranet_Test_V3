<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="RetireFreeTicketList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.RetireFreeTicketList" %>

<asp:Content ID="RetireFreeTicketListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">退休人員乘車證維護作業</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbEmpNo_Search" runat="server" CssClass="text-Right-Black" Text="員工編號：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eEmpNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eEmpName_Search" runat="server" CssClass="text-Left-Black" Width="50%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbRetireYear_Search" runat="server" CssClass="text-Right-Black" Text="退休年度：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eRetireYear_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbTicketYear_Search" runat="server" CssClass="text-Right-Black" Text="乘車證年度：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eTicketYear_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbSerach" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSerach_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
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
    <asp:Panel ID="plShowData_Main" runat="server" CssClass="ShowPanel">
        <div class="ShowPanel-Detail_C">
            <asp:GridView ID="gridShowData_Main" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="EmpNo" DataSourceID="sdsShowData_MainList" ForeColor="#333333" GridLines="None" PageSize="5" Width="100%">
                <AlternatingRowStyle BackColor="White" />
                <Columns>
                    <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                    <asp:BoundField DataField="EmpNo" HeaderText="員工編號" ReadOnly="True" SortExpression="EmpNo" />
                    <asp:BoundField DataField="Name" HeaderText="姓名" SortExpression="Name" />
                    <asp:BoundField DataField="AssumeDate" DataFormatString="{0:D}" HeaderText="AssumeDate" SortExpression="AssumeDate" Visible="False" />
                    <asp:BoundField DataField="RetireDate" DataFormatString="{0:D}" HeaderText="退休日期" SortExpression="RetireDate" />
                    <asp:BoundField DataField="TicketType" HeaderText="TicketType" SortExpression="TicketType" Visible="False" />
                    <asp:BoundField DataField="TicketType_C" HeaderText="優待類別" ReadOnly="True" SortExpression="TicketType_C" />
                    <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" />
                    <asp:BoundField DataField="BuMan" HeaderText="BuMan" SortExpression="BuMan" Visible="False" />
                    <asp:BoundField DataField="BuManName" HeaderText="建檔人" SortExpression="BuManName" />
                    <asp:BoundField DataField="BuDate" DataFormatString="{0:D}" HeaderText="建檔日期" SortExpression="BuDate" />
                    <asp:BoundField DataField="ModifyMan" HeaderText="ModifyMan" SortExpression="ModifyMan" Visible="False" />
                    <asp:BoundField DataField="ModifyManName" HeaderText="ModifyMan" SortExpression="ModifyManName" Visible="False" />
                    <asp:BoundField DataField="ModifyDate" HeaderText="ModifyDate" SortExpression="ModifyDate" Visible="False" />
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
        </div>
        <div class="ShowPanel-Detail_B">
            <asp:FormView ID="fvRetireFreeTicketA" runat="server" DataKeyNames="EmpNo" DataSourceID="sdsShowData_Main" Width="100%" OnDataBound="fvRetireFreeTicketA_DataBound">
                <EditItemTemplate>
                    <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Black" OnClick="bbOK_Edit_Click" Text="確定" Width="120px" />
                    &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                    <asp:UpdatePanel ID="upDataEdit" runat="server">
                        <ContentTemplate>
                            <table class="TableSetting">
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbEmpNo_Edit" runat="server" CssClass="text-Right-Blue" Text="員工編號：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eEmpNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                        <asp:Label ID="eEmpName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Name") %>' Width="90%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbAssumeDate_Edit" runat="server" CssClass="text-Right-Blue" Text="到職日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eAssumeDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssumeDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbRetireDate_Edit" runat="server" CssClass="text-Right-Blue" Text="退休日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eRetireDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RetireDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbTicketType_Edit" runat="server" CssClass="text-Right-Blue" Text="優待類別：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="4">
                                        <asp:DropDownList ID="ddlTicketType_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="ddlTicketType_Edit_TextChanged" />
                                        <asp:Label ID="eTicketType_Edit" runat="server" Text='<%# Eval("TicketType") %>' Visible="false" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColBorder ColWidth-5Col MultiLine_Low">
                                        <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                    </td>
                                    <td class="ColBorder ColWidth-5Col MultiLine_Low" colspan="4">
                                        <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="97%" Height="95%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                        <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                        <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="50%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                        <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                        <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="50%" />
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
                        </ContentTemplate>
                    </asp:UpdatePanel>
                </EditItemTemplate>
                <InsertItemTemplate>
                    <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Black" OnClick="bbOK_INS_Click" Text="確定" Width="120px" />
                    &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                    <asp:UpdatePanel ID="upDataInsert" runat="server">
                        <ContentTemplate>
                            <table class="TableSetting">
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbEmpNo_INS" runat="server" CssClass="text-Right-Blue" Text="員工編號：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:TextBox ID="eEmpNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_INS_TextChanged" Text='<%# Eval("EmpNo") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                        <asp:Label ID="eEmpName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Name") %>' Width="90%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbAssumeDate_INS" runat="server" CssClass="text-Right-Blue" Text="到職日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eAssumeDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssumeDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbRetireDate_INS" runat="server" CssClass="text-Right-Blue" Text="退休日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eRetireDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RetireDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbTicketType_INS" runat="server" CssClass="text-Right-Blue" Text="優待類別：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="4">
                                        <asp:DropDownList ID="ddlTicketType_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="ddlTicketType_INS_TextChanged" />
                                        <asp:Label ID="eTicketType_INS" runat="server" Text='<%# Eval("TicketType") %>' Visible="false" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColBorder ColWidth-5Col MultiLine_Low">
                                        <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                    </td>
                                    <td class="ColBorder ColWidth-5Col MultiLine_Low" colspan="4">
                                        <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="97%" Height="95%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                        <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                        <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="50%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                        <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                        <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="50%" />
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
                        </ContentTemplate>
                    </asp:UpdatePanel>
                </InsertItemTemplate>
                <EmptyDataTemplate>
                    <asp:Button ID="bbNew_Empty" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增退休員工" Width="120px" />
                </EmptyDataTemplate>
                <ItemTemplate>
                    <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增退休員工" Width="120px" />
                    &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CssClass="button-Blue" CommandName="Edit" Text="修改員工資料" Width="120px" />
                    &nbsp;<asp:Button ID="bbDel_List" runat="server" CausesValidation="False" CssClass="button-Red" OnClick="bbDel_List_Click" Text="刪除員工資料" Width="120px" Visible="false" />
                    <table class="TableSetting">
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbEmpNo_List" runat="server" CssClass="text-Right-Blue" Text="員工編號：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="eEmpNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                <asp:Label ID="eEmpName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Name") %>' Width="90%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbAssumeDate_List" runat="server" CssClass="text-Right-Blue" Text="到職日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="eAssumeDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssumeDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbRetireDate_List" runat="server" CssClass="text-Right-Blue" Text="退休日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="eRetireDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RetireDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbTicketType_List" runat="server" CssClass="text-Right-Blue" Text="優待類別：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col" colspan="4">
                                <asp:Label ID="eTicketType_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TicketType_C") %>' Width="90%" />
                                <asp:Label ID="eTicketType_List" runat="server" Text='<%# Eval("TicketType") %>' Visible="false" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColBorder ColWidth-5Col MultiLine_Low">
                                <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                            </td>
                            <td class="ColBorder ColWidth-5Col MultiLine_Low" colspan="4">
                                <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Width="97%" Height="95%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="50%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="50%" />
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
        </div>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsShowData_MainList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT EmpNo, Name, AssumeDate, RetireDate, TicketType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '退休人員乘車證  RetireFreeTicketTicketType') AND (CLASSNO = a.TicketType)) AS TicketType_C, Remark, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuManName, BuDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName, ModifyDate FROM RetireFreeTicketA AS a WHERE (ISNULL(EmpNo, '') = '')"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsShowData_Main" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" DeleteCommand="DELETE FROM RetireFreeTicketA WHERE (EmpNo = @EmpNo)" InsertCommand="INSERT INTO RetireFreeTicketA(EmpNo, Name, AssumeDate, RetireDate, TicketType, Remark, BuMan, BuDate) VALUES (@EmpNo, @Name, @AssumeDate, @RetireDate, @TicketType, @Remark, @BuMan, @BuDate)" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT EmpNo, Name, AssumeDate, RetireDate, TicketType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '退休人員乘車證  RetireFreeTicketTicketType') AND (CLASSNO = a.TicketType)) AS TicketType_C, Remark, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuManName, BuDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName, ModifyDate FROM RetireFreeTicketA AS a WHERE (EmpNo = @EmpNo)" UpdateCommand="UPDATE RetireFreeTicketA SET TicketType = @TicketType, Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate WHERE (EmpNo = @EmpNo)">
        <DeleteParameters>
            <asp:Parameter Name="EmpNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="EmpNo" />
            <asp:Parameter Name="Name" />
            <asp:Parameter Name="AssumeDate" />
            <asp:Parameter Name="RetireDate" />
            <asp:Parameter Name="TicketType" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="BuDate" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridShowData_Main" Name="EmpNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="TicketType" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="EmpNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <br />
    <br />
    <asp:Panel ID="plShowData_Detail" runat="server" CssClass="ShowPanel-Detail">
        <div class="ShowPanel-Detail_C">
            <asp:GridView ID="gridShowData_Detail" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="IndexNo" DataSourceID="sdsShowData_DetailList" ForeColor="#333333" GridLines="None" PageSize="5" Width="100%">
                <AlternatingRowStyle BackColor="White" />
                <Columns>
                    <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                    <asp:BoundField DataField="IndexNo" HeaderText="IndexNo" ReadOnly="True" SortExpression="IndexNo" Visible="False" />
                    <asp:BoundField DataField="EmpNo" HeaderText="EmpNo" SortExpression="EmpNo" Visible="False" />
                    <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                    <asp:BoundField DataField="AssignDate" DataFormatString="{0:D}" HeaderText="申辦日期" SortExpression="AssignDate" />
                    <asp:BoundField DataField="TicketYear" HeaderText="適用年度" SortExpression="TicketYear" />
                    <asp:BoundField DataField="RemarkB" HeaderText="備註" SortExpression="RemarkB" />
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
        </div>
        <div class="ShowPanel-Detail_B">
            <asp:FormView ID="fvRetireFreeTicketB" runat="server" DataKeyNames="IndexNo" DataSourceID="sdsShowData_Detail" Width="100%" OnDataBound="fvRetireFreeTicketB_DataBound">
                <EditItemTemplate>
                    <asp:Button ID="bbDetailOK_EditB" runat="server" CssClass="button-Black" CausesValidation="True" OnClick="bbDetailOK_EditB_Click" Text="確定" Width="120px" />
                    &nbsp;<asp:Button ID="bbDetailCancel_EditB" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                    <asp:UpdatePanel ID="upDataEdit" runat="server">
                        <ContentTemplate>
                            <table class="TableSetting">
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbItems_EditB" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                                        <asp:Label ID="eIndexNo_EditB" runat="server" Text='<%# Eval("IndexNo") %>' Visible="false" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eItems_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbEmpNo_EditB" runat="server" CssClass="text-Right-Blue" Text="員工編號：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                        <asp:Label ID="eEmpNo_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="35%" />
                                        <asp:Label ID="eEmpName_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="50%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbTicketYear_EditB" runat="server" CssClass="text-Right-Blue" Text="適用年度：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:TextBox ID="eTicketYear_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TicketYear") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColWidth-5Col"></td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbAssignDate_EditB" runat="server" CssClass="text-Right-Blue" Text="申辦日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:TextBox ID="eAssignDate_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColBorder ColWidth-5Col MultiLine_Low">
                                        <asp:Label ID="lbRemarkB_EditB" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                    </td>
                                    <td class="ColBorder ColWidth-5Col MultiLine_Low" colspan="4">
                                        <asp:TextBox ID="eRemarkB_EditB" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Width="97%" Height="95%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbBuDate_EditB" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eBuDate_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbBuMan_EditB" runat="server" CssClass="text-Left-Black" Text="建檔人：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                        <asp:Label ID="eBuMan_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                        <asp:Label ID="eBuManName_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="50%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbModifyDate_EditB" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eModifyDate_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbModifyMan_EditB" runat="server" CssClass="text-Left-Black" Text="異動人：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                        <asp:Label ID="eModifyMan_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                        <asp:Label ID="eModifyManName_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="50%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColWidth-5Col"></td>
                                    <td class="ColWidth-5Col"></td>
                                    <td class="ColWidth-5Col"></td>
                                    <td class="ColWidth-5Col"></td>
                                    <td class="ColWidth-5Col"></td>
                                </tr>
                            </table>
                        </ContentTemplate>
                    </asp:UpdatePanel>
                </EditItemTemplate>
                <InsertItemTemplate>
                    <asp:Button ID="bbDetailOK_INSB" runat="server" CssClass="button-Black" CausesValidation="True" OnClick="bbDetailOK_INSB_Click" Text="確定" Width="120px" />
                    &nbsp;<asp:Button ID="bbDetailCancel_INSB" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                    <asp:UpdatePanel ID="upDataInsert" runat="server">
                        <ContentTemplate>
                            <table class="TableSetting">
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbItems_INSB" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                                        <asp:Label ID="eIndexNo_INSB" runat="server" Text='<%# Eval("IndexNo") %>' Visible="false" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eItems_INSB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbEmpNo_INSB" runat="server" CssClass="text-Right-Blue" Text="員工編號：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                        <asp:Label ID="eEmpNo_INSB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="35%" />
                                        <asp:Label ID="eEmpName_INSB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="50%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbTicketYear_INSB" runat="server" CssClass="text-Right-Blue" Text="適用年度：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:TextBox ID="eTicketYear_INSB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TicketYear") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColWidth-5Col"></td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbAssignDate_INSB" runat="server" CssClass="text-Right-Blue" Text="申辦日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:TextBox ID="eAssignDate_INSB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColBorder ColWidth-5Col MultiLine_Low">
                                        <asp:Label ID="lbRemarkB_INSB" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                    </td>
                                    <td class="ColBorder ColWidth-5Col MultiLine_Low" colspan="4">
                                        <asp:TextBox ID="eRemarkB_INSB" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RemarkB") %>' Width="97%" Height="95%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbBuDate_INSB" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eBuDate_INSB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbBuMan_INSB" runat="server" CssClass="text-Left-Black" Text="建檔人：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                        <asp:Label ID="eBuMan_INSB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                        <asp:Label ID="eBuManName_INSB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="50%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbModifyDate_INSB" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="eModifyDate_INSB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col">
                                        <asp:Label ID="lbModifyMan_INSB" runat="server" CssClass="text-Left-Black" Text="異動人：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                        <asp:Label ID="eModifyMan_INSB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                        <asp:Label ID="eModifyManName_INSB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="50%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColWidth-5Col"></td>
                                    <td class="ColWidth-5Col"></td>
                                    <td class="ColWidth-5Col"></td>
                                    <td class="ColWidth-5Col"></td>
                                    <td class="ColWidth-5Col"></td>
                                </tr>
                            </table>
                        </ContentTemplate>
                    </asp:UpdatePanel>
                </InsertItemTemplate>
                <EmptyDataTemplate>
                    <asp:Button ID="bbDetailNew_Empty" runat="server" CssClass="button-Blue" CausesValidation="false" CommandName="New" Text="新增年度乘車證" Width="120px" />
                </EmptyDataTemplate>
                <ItemTemplate>
                    <asp:Button ID="bbDetailNew_ListB" runat="server" CssClass="button-Black" CausesValidation="False" CommandName="New" Text="新增年度乘車證" Width="120px" />
                    &nbsp;<asp:Button ID="bbDetailEdit_ListB" runat="server" CssClass="button-Blue" CausesValidation="False" CommandName="Edit" Text="修改乘車證資料" Width="120px" />
                    &nbsp;<asp:Button ID="bbDetailDel_ListB" runat="server" CssClass="button-Red" CausesValidation="False" OnClick="bbDetailDel_ListB_Click" Text="刪除乘車證資料" Width="120px" Visible="false" />
                    <table class="TableSetting">
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbItems_ListB" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                                <asp:Label ID="eIndexNo_ListB" runat="server" Text='<%# Eval("IndexNo") %>' Visible="false" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="eItems_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbEmpNo_ListB" runat="server" CssClass="text-Right-Blue" Text="員工編號：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                <asp:Label ID="eEmpNo_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="35%" />
                                <asp:Label ID="eEmpName_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="50%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbTicketYear_ListB" runat="server" CssClass="text-Right-Blue" Text="適用年度：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="eTicketYear_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TicketYear") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColWidth-5Col"></td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbAssignDate_ListB" runat="server" CssClass="text-Right-Blue" Text="申辦日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="eAssignDate_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColBorder ColWidth-5Col MultiLine_Low">
                                <asp:Label ID="lbRemarkB_ListB" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                            </td>
                            <td class="ColBorder ColWidth-5Col MultiLine_Low" colspan="4">
                                <asp:TextBox ID="eRemarkB_ListB" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("RemarkB") %>' Width="97%" Height="95%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbBuDate_ListB" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="eBuDate_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbBuMan_ListB" runat="server" CssClass="text-Left-Black" Text="建檔人：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                <asp:Label ID="eBuMan_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                <asp:Label ID="eBuManName_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="50%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbModifyDate_ListB" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="eModifyDate_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col">
                                <asp:Label ID="lbModifyMan_ListB" runat="server" CssClass="text-Left-Black" Text="異動人：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-5Col" colspan="2">
                                <asp:Label ID="eModifyMan_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                <asp:Label ID="eModifyManName_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="50%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColWidth-5Col"></td>
                            <td class="ColWidth-5Col"></td>
                            <td class="ColWidth-5Col"></td>
                            <td class="ColWidth-5Col"></td>
                            <td class="ColWidth-5Col"></td>
                        </tr>
                    </table>
                </ItemTemplate>
            </asp:FormView>
        </div>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsShowData_DetailList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT IndexNo, EmpNo, Items, AssignDate, TicketYear, RemarkB FROM RetireFreeTicketB WHERE (EmpNo = @EmpNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridShowData_Main" Name="EmpNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsShowData_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" DeleteCommand="DELETE FROM RetireFreeTicketB WHERE (IndexNo = @IndexNo)" InsertCommand="INSERT INTO RetireFreeTicketB(IndexNo, EmpNo, Items, AssignDate, TicketYear, RemarkB, BuMan, BuDate) VALUES (@IndexNo, @EmpNo, @Items, @AssignDate, @TicketYear, @RemarkB, @BuMan, @BuDate)" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT IndexNo, EmpNo, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = b.EmpNo)) AS EmpName, Items, AssignDate, TicketYear, RemarkB, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = b.BuMan)) AS BuManName, BuDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = b.ModifyMan)) AS ModifyManName, ModifyDate FROM RetireFreeTicketB AS b WHERE (IndexNo = @IndexNo)" UpdateCommand="UPDATE RetireFreeTicketB SET AssignDate = @AssignDate, TicketYear = @TicketYear, RemarkB = @RemarkB, ModifyDate = @ModifyDate, ModifyMan = @ModifyMan WHERE (IndexNo = @IndexNo)">
        <DeleteParameters>
            <asp:Parameter Name="IndexNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="IndexNo" />
            <asp:Parameter Name="EmpNo" />
            <asp:Parameter Name="Items" />
            <asp:Parameter Name="AssignDate" />
            <asp:Parameter Name="TicketYear" />
            <asp:Parameter Name="RemarkB" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="BuDate" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridShowData_Detail" Name="IndexNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="AssignDate" />
            <asp:Parameter Name="TicketYear" />
            <asp:Parameter Name="RemarkB" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="IndexNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
        <!--預留做為以後要做報表時使用-->
    </asp:Panel>
</asp:Content>
