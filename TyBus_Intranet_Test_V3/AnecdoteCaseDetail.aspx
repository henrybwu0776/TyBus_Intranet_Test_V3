<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="AnecdoteCaseDetail.aspx.cs" Inherits="TyBus_Intranet_Test_V3.AnecdoteCaseDetail" %>

<asp:Content ID="AnecdoteCaseDetailForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">肇事案件處理歷程</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:CheckBox ID="chkHasInsurance" runat="server" CssClass="text-Left-Black" Text="已出險" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Check" runat="server" CssClass="text-Right-Blue" Text="站別 (編號)：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Start_Check" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="eDepNo_Split_Check" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_End_Check" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCaseDate_Check" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eCaseDate_Start_Check" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="eCaseDate_Split_Check" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eCaseDate_End_Check" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbCarID_Search" runat="server" CssClass="text-Right-Blue" Text="牌照號碼" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eCarID_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbDriver_Search" runat="server" CssClass="text-Right-Blue" Text="駕駛員" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eDriver_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eDriverName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" Text="預覽" CssClass="button-Black" runat="server" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbCancel" Text="結束" CssClass="button-Red" runat="server" OnClick="bbCancel_Click" Width="90%" />
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
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plMainDataShow" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridAnecdoteCaseA_List" runat="server" AllowPaging="True" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataSourceID="sdsAnecdoteCaseA_List" GridLines="None" PageSize="5" Width="100%" AutoGenerateColumns="False" DataKeyNames="CaseNo">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="CaseNo" HeaderText="序號" ReadOnly="True" SortExpression="CaseNo" />
                <asp:CheckBoxField DataField="HasInsurance" HeaderText="出險" SortExpression="HasInsurance" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="站別" SortExpression="DepName" />
                <asp:BoundField DataField="BuildDate" DataFormatString="{0:d}" HeaderText="發生日期" SortExpression="BuildDate" />
                <asp:BoundField DataField="BuildMan" HeaderText="BuildMan" SortExpression="BuildMan" Visible="False" />
                <asp:BoundField DataField="BuildManName" HeaderText="建檔人" ReadOnly="True" SortExpression="BuildManName" />
                <asp:BoundField DataField="Car_ID" HeaderText="車號" SortExpression="Car_ID" />
                <asp:BoundField DataField="Driver" HeaderText="Driver" SortExpression="Driver" Visible="False" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員" SortExpression="DriverName" />
                <asp:BoundField DataField="InsuMan" HeaderText="保險經辦人" SortExpression="InsuMan" />
                <asp:BoundField DataField="AnecdotalResRatio" HeaderText="肇責比率" SortExpression="AnecdotalResRatio" />
                <asp:CheckBoxField DataField="IsNoDeduction" HeaderText="免扣精勤" SortExpression="IsNoDeduction" />
                <asp:BoundField DataField="DeductionDate" DataFormatString="{0:d}" HeaderText="扣發日期" SortExpression="DeductionDate" />
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
        <br />
        <asp:GridView ID="gridAnecdoteCaseDetailList" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="CaseNoItem"
            DataSourceID="sdsAnecdoteCaseC_List" ForeColor="#333333" GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="CaseNoItem" HeaderText="CaseNoItem" ReadOnly="True" SortExpression="CaseNoItem" Visible="False" />
                <asp:BoundField DataField="CaseNo" HeaderText="CaseNo" SortExpression="CaseNo" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                <asp:BoundField DataField="ContactDate" DataFormatString="{0:d}" HeaderText="連絡日期" SortExpression="ContactDate" />
                <asp:BoundField DataField="Excutort_C" HeaderText="處理人員" ReadOnly="True" SortExpression="Excutort_C" />
                <asp:BoundField DataField="ContactPerson" HeaderText="連絡對象" SortExpression="ContactPerson" />
                <asp:BoundField DataField="ContactNote" HeaderText="處理概述" SortExpression="ContactNote" />
                <asp:BoundField DataField="AssignDate" DataFormatString="{0:d}" HeaderText="轉交日期" SortExpression="AssignDate" />
                <asp:BoundField DataField="AssignedMan_C" HeaderText="轉交人員" ReadOnly="True" SortExpression="AssignedMan_C" />
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
        <asp:FormView ID="fvAnecdoteCaseC" runat="server" DataKeyNames="CaseNoItem" DataSourceID="sdsAnecdoteCaseC_Data" Width="100%" OnDataBound="fvAnecdoteCaseC_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Update" Text="確定" UseSubmitBehavior="False" OnClientClick="this.disabled=true;" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseNo_Edit" runat="server" CssClass="text-Right-Blue" Text="肇事案號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="90%" />
                                    <asp:Label ID="eCaseNoItem_Edit" runat="server" Text='<%# Eval("CaseNoItem") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbItems_Edit" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="90%" />
                                    <br />
                                    <asp:Label ID="eBuildMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan_C") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbExcutort_Edit" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eExcutort_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eExcutort_Edit_TextChanged" Text='<%# Bind("Excutort") %>' Width="35%" />
                                    <asp:Label ID="eExcutort_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Excutort_C") %>' Width="55%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" rowspan="3">
                                    <asp:Label ID="lbContactNote_Edit" runat="server" CssClass="text-Right-Blue" Text="處理概述：" Width="90%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="3" rowspan="3">
                                    <asp:TextBox ID="eContactNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("ContactNote") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbContextDate_Edit" runat="server" CssClass="text-Right-Blue" Text="連絡日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eContactDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContactDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbContactPerson_Edit" runat="server" CssClass="text-Right-Blue" Text="連絡對象：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eContactPerson_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ContactPerson") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignDate_Edit" runat="server" CssClass="text-Right-Blue" Text="轉交日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Assigndate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignedMan_Edit" runat="server" CssClass="text-Right-Blue" Text="轉交人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignedMan_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignedMan_Edit_TextChanged" Text='<%# Bind("AssignedMan") %>' Width="90%" />
                                    <br />
                                    <asp:Label ID="eAssignedMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignedMan_C") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="90%" />
                                    <br />
                                    <asp:Label ID="eModifyMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="90%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" rowspan="2">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="異動說明：" Width="90%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="5" rowspan="2">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
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
                <asp:Button ID="bbNew_Empty" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="NEW" Text="新增" />
            </EmptyDataTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Insert" Text="確定" UseSubmitBehavior="False" OnClientClick="this.disabled=true;" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseNo_INS" runat="server" CssClass="text-Right-Blue" Text="肇事案號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_INS" runat="server" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbItems_INS" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eItems_INS" runat="server" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildDate_INS" runat="server" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildMan_INS" runat="server" CssClass="text-Left-Black" Width="90%" />
                                    <br />
                                    <asp:Label ID="eBuildMan_C_INS" runat="server" CssClass="text-Left-Black" Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbExcutort_INS" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eExcutort_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eExcutort_INS_TextChanged" Text='<%# Bind("Excutort") %>' Width="35%" />
                                    <asp:Label ID="eExcutort_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Excutort_C") %>' Width="55%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" rowspan="3">
                                    <asp:Label ID="lbContactNote_INS" runat="server" CssClass="text-Right-Blue" Text="處理概述：" Width="90%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="3" rowspan="3">
                                    <asp:TextBox ID="eContactNote_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("ContactNote") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbContextDate_INS" runat="server" CssClass="text-Right-Blue" Text="連絡日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eContactDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContactDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbContactPerson_INS" runat="server" CssClass="text-Right-Blue" Text="連絡對象：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eContactPerson_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ContactPerson") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignDate_INS" runat="server" CssClass="text-Right-Blue" Text="轉交日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Assigndate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignedMan_INS" runat="server" CssClass="text-Right-Blue" Text="轉交人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignedMan_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignedMan_INS_TextChanged" Text='<%# Bind("AssignedMan") %>' Width="90%" />
                                    <br />
                                    <asp:Label ID="eAssignedMan_C_INS" runat="server" CssClass="text-Left-Black" Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="90%" />
                                    <br />
                                    <asp:Label ID="eModifyMan_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="90%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" rowspan="2">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="異動說明：" Width="90%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="5" rowspan="2">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Enabled="false" Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
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
                <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="Edit" Text="編輯" />
                &nbsp;<asp:Button ID="bbDelete_List" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Delete" Text="刪除" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseNo_List" runat="server" CssClass="text-Right-Blue" Text="肇事案號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseNoItem_List" runat="server" Text='<%# Eval("CaseNoItem") %>' Visible="false" />
                            <asp:Label ID="eCaseNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItems_List" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuildDate_LIst" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuildMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="90%" />
                            <br />
                            <asp:Label ID="eBuildMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan_C") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbExcutort_List" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eExcutort_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Excutort") %>' Width="35%" />
                            <asp:Label ID="eExcutort_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Excutort_C") %>' Width="55%" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine-Low" rowspan="3">
                            <asp:Label ID="lbContactNote_List" runat="server" CssClass="text-Right-Blue" Text="處理概述：" Width="90%" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="3" rowspan="3">
                            <asp:TextBox ID="eContactNote_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("ContactNote") %>' Enabled="false" Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbContextDate_List" runat="server" CssClass="text-Right-Blue" Text="連絡日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eContactDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContactDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbContactPerson_List" runat="server" CssClass="text-Right-Blue" Text="連絡對象：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eContactPerson_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContactPerson") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignDate_List" runat="server" CssClass="text-Right-Blue" Text="轉交日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAssignDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Assigndate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignedMan_List" runat="server" CssClass="text-Right-Blue" Text="轉交人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAssignedMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignedMan") %>' Width="90%" />
                            <br />
                            <asp:Label ID="eAssignedMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignedMan_C") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="90%" />
                            <br />
                            <asp:Label ID="eModifyMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="90%" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine-Low" rowspan="2">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="異動說明：" Width="90%" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="5" rowspan="2">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Enabled="false" Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
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
    </asp:Panel>
    <asp:SqlDataSource ID="sdsAnecdoteCaseA_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = AnecdoteCase.BuildMan)) AS BuildManName, Car_ID, Driver, DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, CaseOccurrence FROM AnecdoteCase WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsAnecdoteCaseC_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNoItem, CaseNo, Items, ContactPerson, ContactNote, ContactDate, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = B.Excutort)) AS Excutort_C, AssignDate, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (EMPNO = B.AssignedMan)) AS AssignedMan_C FROM AnecdoteCaseC AS B WHERE (CaseNo = @CaseNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridAnecdoteCaseA_List" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>

    <asp:SqlDataSource ID="sdsAnecdoteCaseC_Data" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        DeleteCommand="DELETE FROM AnecdoteCaseC WHERE (CaseNoItem = @CaseNoItem)"
        InsertCommand="INSERT INTO AnecdoteCaseC(CaseNoItem, CaseNo, Items, ContactPerson, ContactDate, ContactNote, Excutort, AssignDate, AssignedMan, BuildDate, BuildMan) VALUES (@CaseNoItem, @CaseNo, @Items, @ContactPerson, @ContactDate, @ContactNote, @Excutort, @AssignDate, @AssignedMan, @BuildDate, @BuildMan)"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNoItem, CaseNo, Items, ContactPerson, ContactNote, ContactDate, Excutort, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = B.Excutort)) AS Excutort_C, AssignDate, AssignedMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (EMPNO = B.AssignedMan)) AS AssignedMan_C, BuildDate, BuildMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = B.BuildMan)) AS BuildMan_C, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = B.ModifyMan)) AS ModifyMan_C, Remark FROM AnecdoteCaseC AS B WHERE (CaseNoItem = @CaseNoItem)"
        UpdateCommand="UPDATE AnecdoteCaseC SET ContactPerson = @ContactPerson, ContactNote = @ContactNote, ContactDate = @ContactDate, Excutort = @Excutort, AssignDate = @AssignDate, AssignedMan = @AssignedMan, ModifyDate = @ModifyDate, ModifyMan = @ModifyMan, Remark = @Remark WHERE (CaseNoItem = @CaseNoItem)"
        OnDeleted="sdsAnecdoteCaseC_Data_Deleted"
        OnInserted="sdsAnecdoteCaseC_Data_Inserted"
        OnInserting="sdsAnecdoteCaseC_Data_Inserting"
        OnUpdated="sdsAnecdoteCaseC_Data_Updated"
        OnUpdating="sdsAnecdoteCaseC_Data_Updating">
        <DeleteParameters>
            <asp:Parameter Name="CaseNoItem" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="CaseNoItem" />
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="Items" />
            <asp:Parameter Name="ContactPerson" />
            <asp:Parameter Name="ContactDate" />
            <asp:Parameter Name="ContactNote" />
            <asp:Parameter Name="Excutort" />
            <asp:Parameter Name="AssignDate" />
            <asp:Parameter Name="AssignedMan" />
            <asp:Parameter Name="BuildDate" />
            <asp:Parameter Name="BuildMan" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridAnecdoteCaseDetailList" Name="CaseNoItem" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="ContactPerson" />
            <asp:Parameter Name="ContactNote" />
            <asp:Parameter Name="ContactDate" />
            <asp:Parameter Name="Excutort" />
            <asp:Parameter Name="AssignDate" />
            <asp:Parameter Name="AssignedMan" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="CaseNoItem" />
        </UpdateParameters>
    </asp:SqlDataSource>
</asp:Content>
