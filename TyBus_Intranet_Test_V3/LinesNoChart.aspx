<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="LinesNoChart.aspx.cs" Inherits="TyBus_Intranet_Test_V3.LinesNoChart" %>

<asp:Content ID="LinesNoChartForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">路線代號對照表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbTicketLinesNo_Search" runat="server" CssClass="text-Right-Blue" Text="驗票機路線代號：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eTicketLinesNo_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbERPLinesNo_Search" runat="server" CssClass="text-Right-Blue" Text="ERP 路線代號：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eERPLinesNo_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbGOVLinesNo_Search" runat="server" CssClass="text-Right-Blue" Text="公總路線代號：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eGOVLinesNo_Search" runat="server" CssClass="text-Left-Black" Width="60%" />
                    <asp:TextBox ID="eGOVLinesExtNo_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-7Col" />
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-7Col" colspan="3">
                    <asp:FileUpload ID="fuExcel" runat="server" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯入資料" OnClick="bbExcel_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="95%" />
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
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridLinesNoChartList" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="IndexNo" DataSourceID="sdsLinesNoChart_List" GridLines="None">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="IndexNo" HeaderText="IndexNo" ReadOnly="True" SortExpression="IndexNo" Visible="False" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="主責單位" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="TicketLinesNo" HeaderText="驗票機路線代號" SortExpression="TicketLinesNo" />
                <asp:BoundField DataField="ERPLinesNo" HeaderText="ERP路線代號" SortExpression="ERPLinesNo" />
                <asp:BoundField DataField="GOVLinesNo" HeaderText="公總路線代號" SortExpression="GOVLinesNo" />
                <asp:BoundField DataField="GOVLinesExtNo" HeaderText="特殊營運型態" SortExpression="GOVLinesExtNo" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" />
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
        <asp:FormView ID="fvLinesNoChartDetail" runat="server" Width="100%" DataKeyNames="IndexNo" DataSourceID="sdsLinesNoChart_Detail">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Black" OnClick="bbOK_Edit_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbIndexNo_Edit" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eIndexNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("IndexNo") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="主責單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Text='<%# Eval("DepNo") %>' Width="30%" />
                                    <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbTicketLinesNo_Edit" runat="server" CssClass="text-Right-Blue" Text="驗票機路線代號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eTicketLinesNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TicketLinesNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbERPLinesNo_Edit" runat="server" CssClass="text-Right-Blue" Text="ERP路線代號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eERPLinesNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ERPLinesNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbGOVLinesNo_Edit" runat="server" CssClass="text-Right-Blue" Text="公總路線代號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eGOVLinesNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GOVLinesNo") %>' Width="60%" />
                                    <asp:TextBox ID="eGOVLinesExtNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GOVLinesExtNo") %>' Width="30%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_High ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-6Col" colspan="5">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-6Col" />
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                                    <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-6Col" />
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="最後異動人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="30%" />
                                    <asp:Label ID="eModifyManNAme_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="最後異動日：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CommandName="New" CssClass="button-Blue" Text="新增" Width="120px" />
            </EmptyDataTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_INS_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbIndexNo_INS" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eIndexNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("IndexNo") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="主責單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Eval("DepNo") %>' Width="30%" />
                                    <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbTicketLinesNo_INS" runat="server" CssClass="text-Right-Blue" Text="驗票機路線代號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eTicketLinesNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TicketLinesNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbERPLinesNo_INS" runat="server" CssClass="text-Right-Blue" Text="ERP路線代號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eERPLinesNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ERPLinesNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbGOVLinesNo_INS" runat="server" CssClass="text-Right-Blue" Text="公總路線代號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eGOVLinesNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GOVLinesNo") %>' Width="60%" />
                                    <asp:TextBox ID="eGOVLinesExtNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GOVLinesExtNo") %>' Width="30%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_High ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-6Col" colspan="5">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-6Col" />
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                                    <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-6Col" />
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="最後異動人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="30%" />
                                    <asp:Label ID="eModifyManNAme_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="最後異動日：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </InsertItemTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CssClass="button-Blue" CommandName="Edit" Text="編輯" Width="120px" />
                &nbsp;<asp:Button ID="bbDelete_List" runat="server" CausesValidation="False" CssClass="button-Red" OnClick="bbDelete_List_Click" Text="刪除" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbIndexNo_List" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eIndexNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("IndexNo") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="主責單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="30%" />
                                    <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbTicketLinesNo_List" runat="server" CssClass="text-Right-Blue" Text="驗票機路線代號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eTicketLinesNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TicketLinesNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbERPLinesNo_List" runat="server" CssClass="text-Right-Blue" Text="ERP路線代號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eERPLinesNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ERPLinesNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbGOVLinesNo_List" runat="server" CssClass="text-Right-Blue" Text="公總路線代號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eGOVLinesNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GOVLinesNo") %>' Width="60%" />
                                    <asp:TextBox ID="eGOVLinesExtNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GOVLinesExtNo") %>' Width="30%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_High ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-6Col" colspan="5">
                                    <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" Enabled="false" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-6Col" />
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                                    <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-6Col" />
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="最後異動人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="30%" />
                                    <asp:Label ID="eModifyManNAme_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="最後異動日：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsLinesNoChart_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT IndexNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, ERPLinesNo, GOVLinesNo, GOVLinesExtNo, TicketLinesNo, Remark FROM LinesNoChart AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsLinesNoChart_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" DeleteCommand="DELETE FROM LinesNoChart WHERE (IndexNo = @IndexNo)" InsertCommand="INSERT INTO LinesNoChart(IndexNo, DepNo, ERPLinesNo, GOVLinesNo, GOVLinesExtNo, TicketLinesNo, Remark, BuMan, BuDate) VALUES (@IndexNo, @DepNo, @ERPLinesNo, @GOVLinesNo, @GOVLinesExtNo, @TicketLinesNo, @Remark, @BuMan, @BuDate)" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT IndexNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, ERPLinesNo, GOVLinesNo, GOVLinesExtNo, TicketLinesNo, Remark, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuManName, BuDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName, ModifyDate FROM LinesNoChart AS a WHERE (IndexNo = @IndexNo)" UpdateCommand="UPDATE LinesNoChart SET DepNo = @DepNo, ERPLinesNo = @ERPLinesNo, GOVLinesNo = @GOVLinesNo, GOVLinesExtNo = @GOVLinesExtNo, TicketLinesNo = @TicketLinesNo, Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate WHERE (IndexNo = @IndexNo)">
        <DeleteParameters>
            <asp:Parameter Name="IndexNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="IndexNo" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="ERPLinesNo" />
            <asp:Parameter Name="GOVLinesNo" />
            <asp:Parameter Name="GOVLinesExtNo" />
            <asp:Parameter Name="TicketLinesNo" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="BuDate" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridLinesNoChartList" Name="IndexNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="ERPLinesNo" />
            <asp:Parameter Name="GOVLinesNo" />
            <asp:Parameter Name="GOVLinesExtNo" />
            <asp:Parameter Name="TicketLinesNo" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="IndexNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
</asp:Content>
