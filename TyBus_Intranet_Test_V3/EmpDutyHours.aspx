<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmpDutyHours.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmpDutyHours" %>

<asp:Content ID="EmpDutyHoursForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">員工加班補休</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" GroupingText="查詢條件" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="單位編號：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbEmpNo_Search" runat="server" CssClass="text-Right-Blue" Text="員工編號：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eEmpNo_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit2" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eEmpNo_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
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
    <asp:Panel ID="plEmpDataList" runat="server" GroupingText="員工名單" CssClass="ShowPanel">
        <asp:GridView ID="gridEmpDataList" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="EMPNO"
            DataSourceID="sdsEmpDataList" GridLines="None" Width="100%" PageSize="5" CssClass="text-Left-Black">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True">
                    <ItemStyle Width="10%" />
                </asp:CommandField>
                <asp:BoundField DataField="DEPNO" HeaderText="DEPNO" SortExpression="DEPNO" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="部門別" SortExpression="DepName" ReadOnly="True">
                    <ItemStyle Width="25%" />
                </asp:BoundField>
                <asp:BoundField DataField="TITLE" HeaderText="TITLE" SortExpression="TITLE" Visible="False"></asp:BoundField>
                <asp:BoundField DataField="Title_C" HeaderText="職稱" SortExpression="Title_C" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="EMPNO" HeaderText="員工編號" SortExpression="EMPNO" ReadOnly="True">
                    <ItemStyle Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="NAME" HeaderText="姓名" SortExpression="NAME">
                    <ItemStyle Width="30%" />
                </asp:BoundField>
                <asp:BoundField DataField="UseableHours" HeaderText="可換休時數" ReadOnly="True" SortExpression="UseableHours" />
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
    </asp:Panel>
    <asp:Panel ID="plEmpDutyDataList" runat="server" GroupingText="員工加班/補休資料" CssClass="ShowPanel-Detail">
        <asp:GridView ID="gridEmpDutyDataList" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4"
            DataKeyNames="DutyIndex" DataSourceID="sdsEmpDutyDataList" ForeColor="#333333" GridLines="None" PageSize="5" Width="100%" CssClass="text-Left-Black">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="DutyIndex" HeaderText="序號" ReadOnly="True" SortExpression="DutyIndex" Visible="false" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="部門別" SortExpression="DepName" />
                <asp:BoundField DataField="EmpNo" HeaderText="EmpNo" SortExpression="EmpNo" Visible="False" />
                <asp:BoundField DataField="EmpName" HeaderText="員工" SortExpression="EmpName" />
                <asp:BoundField DataField="DutyType_C" HeaderText="類別" SortExpression="DutyType_C" />
                <asp:BoundField DataField="DutyDateStart" DataFormatString="{0:d}" HeaderText="開始日期" SortExpression="DutyDateStart" />
                <asp:BoundField DataField="StartTime" HeaderText="開始時間" SortExpression="StartTime" />
                <asp:BoundField DataField="DutyDateEnd" DataFormatString="{0:d}" HeaderText="結束日期" SortExpression="DutyDateEnd" />
                <asp:BoundField DataField="EndTime" HeaderText="結束時間" SortExpression="EndTime" />
                <asp:BoundField DataField="DutyHours" HeaderText="加班時數" SortExpression="DutyHours" />
                <asp:BoundField DataField="ESCHours" HeaderText="補休時數" SortExpression="ESCHours" />
                <asp:BoundField DataField="UsableHours" HeaderText="可休時數" SortExpression="UsableHours" />
                <asp:BoundField DataField="BuMan" HeaderText="BuMan" SortExpression="BuMan" Visible="False" />
                <asp:BoundField DataField="BuMan_C" HeaderText="BuMan_C" ReadOnly="True" SortExpression="BuMan_C" Visible="False" />
                <asp:BoundField DataField="BuDate" HeaderText="BuDate" SortExpression="BuDate" Visible="False" />
                <asp:BoundField DataField="AssignReason" HeaderText="事由" SortExpression="AssignReason" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" />
                <asp:BoundField DataField="Inspector" HeaderText="Inspector" SortExpression="Inspector" Visible="False" />
                <asp:BoundField DataField="Inspector_C" HeaderText="Inspector_C" ReadOnly="True" SortExpression="Inspector_C" Visible="False" />
                <asp:CheckBoxField DataField="IsAllowed" HeaderText="已核准" SortExpression="IsAllowed" />
                <asp:CheckBoxField DataField="IsRejected" HeaderText="已駁回" SortExpression="IsRejected" />
                <asp:BoundField DataField="RejectReason" HeaderText="駁回原因" SortExpression="RejectReason" />
                <asp:BoundField DataField="HoursRatio" HeaderText="HoursRatio" SortExpression="HoursRatio" Visible="False" />
                <asp:CheckBoxField DataField="IsDisOneHour" HeaderText="扣除午休1小時" SortExpression="IsDisOneHour" />
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
        <asp:FormView ID="fvEmpDutyDataDetail" runat="server" DataKeyNames="DutyIndex" DataSourceID="sdsEmpDutyDataDetail" Width="100%"
            OnDataBound="fvEmpDutyDataDetail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CssClass="button-Black" CausesValidation="True" OnClick="bbOK_Edit_Click" Text="更新" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDutyIndex_Edit" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eDutyIndex_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyIndex") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepartment_Edit" runat="server" CssClass="text-Right-Blue" Text="部門別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="30%" />
                                    <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eEmployee_Edit" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eEmpNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="30%" />
                                    <asp:Label ID="eEmpName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDutyType_Edit" runat="server" CssClass="text-Right-Blue" Text="類別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eDutyType_Edit" runat="server" Text='<%# Eval("DutyType") %>' Visible="false" />
                                    <asp:Label ID="eDutyType_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyType_C") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="cbIsDisOneHour_Edit" runat="server" CssClass="text-Left-Black" Text="扣除午休1小時" />
                                    <asp:Label ID="eIsDisOneHour_Edit" runat="server" Text='<%# Eval("IsDisOneHour") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDutyDateStart_Edit" runat="server" CssClass="text-Right-Blue" Text="起迄日期時間：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eDutyDateStart_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyDateStart", "{0:yyyy/MM/dd}") %>' Width="15%" />
                                    <asp:TextBox ID="eStartTime_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StartTime") %>' AutoPostBack="true"
                                        OnTextChanged="eStartTime_Edit_TextChanged" Width="10%" />
                                    <asp:Label ID="lbSplit1_Edit" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                                    <asp:TextBox ID="eDutyDateEnd_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyDateEnd", "{0:yyyy/MM/dd}") %>' Width="15%" />
                                    <asp:TextBox ID="eEndTime_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EndTime") %>' AutoPostBack="true"
                                        OnTextChanged="eStartTime_Edit_TextChanged" Width="10%" />
                                    <br />
                                    <asp:Label ID="lbTimeError_Edit" runat="server" CssClass="text-Left-Blue" Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDutyHours_Edit" runat="server" CssClass="text-Right-Blue" Text="時數：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:Label ID="lbHours1_Edit" runat="server" CssClass="text-Right-Blue" Text="加班：" Width="15%" />
                                    <asp:Label ID="eDutyHours_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyHours") %>' Width="30%" />
                                    <asp:Label ID="lbHours2_Edit" runat="server" CssClass="text-Right-Blue" Text="補休：" Width="15%" />
                                    <asp:Label ID="eESCHours_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ESCHours") %>' Width="30%" />
                                    <asp:Label ID="eUsableHours_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UsableHours") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignReason_Edit" runat="server" CssClass="text-Right-Blue" Text="申請原由：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eAssignReason_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignReason") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Height="95%" Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbInspector_Edit" runat="server" CssClass="text-Right-Blue" Text="核准主管：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eInspector_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Inspector") %>' OnTextChanged="eInspector_Edit_TextChanged" AutoPostBack="true" Width="30%" />
                                    <asp:Label ID="eInspector_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Inspector_C") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="cbIsAllowed_Edit" runat="server" CssClass="text-Left-Black" Enabled="true" Text="已核准" />
                                    <asp:Label ID="eIsAllowed_Edit" runat="server" Text='<%# Eval("IsAllowed") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="cbIsRejected_Edit" runat="server" CssClass="text-Left-Black" Enabled="true" Text="已駁回" />
                                    <asp:Label ID="eIsRejected_Edit" runat="server" Text='<%# Eval("IsRejected") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColWidth-8Col">
                                    <asp:Label ID="eHoursRatio_Edit" runat="server" Text='<%# Eval("HoursRatio") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRejectedReason_Edit" runat="server" CssClass="text-Right-Blue" Text="駁回原因：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRejectReason_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("RejectReason") %>' Height="95%" Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                                    <asp:Label ID="eBuMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
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
                <asp:Button ID="bbNew_Empty" runat="server" CssClass="button-Black" CommandName="New" Text="新增" />
            </EmptyDataTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_Insert" runat="server" CssClass="button-Black" CausesValidation="True" OnClick="bbOK_Insert_Click" Text="確定" />
                &nbsp;<asp:Button ID="bbCancel_Insert" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDutyIndex_Insert" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDutyIndex_Insert" runat="server" CssClass="text-Left-Black" Enabled="false" Text='<%# Eval("DutyIndex") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepartment_Insert" runat="server" CssClass="text-Right-Blue" Text="部門別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="30%"
                                        OnTextChanged="eDepNo_Insert_TextChanged" AutoPostBack="True" />
                                    <asp:Label ID="eDepName_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eEmployee_Insert" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eEmpNo_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="30%" AutoPostBack="True"
                                        OnTextChanged="eEmpNo_Insert_TextChanged" />
                                    <asp:Label ID="eEmpName_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="55%" />
                                    <br />
                                    <asp:Label ID="EmpDataError_Insert" runat="server" CssClass="text-Left-Blue" Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDutyType_Insert" runat="server" CssClass="text-Right-Blue" Text="類別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlDutyType_Insert" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True"
                                        DataSourceID="sdsDutyType_INS" DataTextField="ClassTxt" DataValueField="ClassNo"
                                        OnSelectedIndexChanged="ddlDutyType_Insert_SelectedIndexChanged" />
                                    <br />
                                    <asp:Label ID="eDutyType_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyType") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="cbIsDisOneHour_Insert" runat="server" CssClass="text-Left-Black" Text="扣除午休1小時" />
                                    <asp:Label ID="eIsDisOneHour_Insert" runat="server" Text='<%# Eval("IsDisOneHour") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDutyDateStart_Insert" runat="server" CssClass="text-Right-Blue" Text="起迄日期時間：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eDutyDateStart_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyDateStart", "{0:yyyy/MM/dd}") %>' Width="15%" />
                                    <asp:TextBox ID="eStartTime_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StartTime") %>' AutoPostBack="true"
                                        OnTextChanged="eStartTime_Insert_TextChanged" Width="10%" />
                                    <asp:Label ID="lbSplit1_Insert" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                                    <asp:TextBox ID="eDutyDateEnd_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyDateEnd", "{0:yyyy/MM/dd}") %>' Width="15%" />
                                    <asp:TextBox ID="eEndTime_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EndTime") %>' AutoPostBack="true"
                                        OnTextChanged="eStartTime_Insert_TextChanged" Width="10%" />
                                    <br />
                                    <asp:Label ID="lbTimeError_Insert" runat="server" CssClass="text-Left-Blue" Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDutyHours_Insert" runat="server" CssClass="text-Right-Blue" Text="時數：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:Label ID="lbHours1_Insert" runat="server" CssClass="text-Right-Blue" Text="加班：" Width="15%" />
                                    <asp:Label ID="eDutyHours_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyHours") %>' Width="30%" />
                                    <asp:Label ID="lbHours2_Insert" runat="server" CssClass="text-Right-Blue" Text="補休：" Width="15%" />
                                    <asp:Label ID="eESCHours_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ESCHours") %>' Width="30%" />
                                    <asp:Label ID="eUsableHours_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UsableHours") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignReason_Insert" runat="server" CssClass="text-Right-Blue" Text="申請原由：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eAssignReason_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignReason") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_Insert" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_Insert" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Height="95%" Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbInspector_Insert" runat="server" CssClass="text-Right-Blue" Text="核准主管：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eInspector_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Inspector") %>' Width="30%" AutoPostBack="True"
                                        OnTextChanged="eInspector_Insert_TextChanged" />
                                    <asp:Label ID="eInspector_C_Insert" runat="server" CssClass="text-Left-Black" Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="cbIsAllowed_Insert" runat="server" CssClass="text-Left-Black" Text="已核准" />
                                    <asp:Label ID="eIsAllowed_Insert" runat="server" Text='<%# Eval("IsAllowed") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="cbIsRejected_Insert" runat="server" CssClass="text-Left-Black" Text="已駁回" />
                                    <asp:Label ID="eIsRejected_Insert" runat="server" Text='<%# Eval("IsRejected") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColWidth-8Col">
                                    <asp:Label ID="eHoursRatio_Insert" runat="server" Text='<%# Eval("HoursRatio") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRejectedReason_Insert" runat="server" CssClass="text-Right-Blue" Text="駁回原因：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRejectReason_Insert" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RejectReason") %>' Height="95%" Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_Insert" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                                    <asp:Label ID="eBuMan_C_Insert" runat="server" CssClass="text-Left-Black" Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_Insert" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDate_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
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
                <asp:Button ID="bbNew_List" runat="server" CssClass="button-Black" CausesValidation="False" CommandName="New" Text="新增" Width="100px" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CssClass="button-Black" CausesValidation="False" CommandName="Edit" Text="編輯" Width="100px" />
                &nbsp;<asp:Button ID="bbDelete_List" runat="server" CssClass="button-Red" CausesValidation="False" Text="刪除" OnClick="bbDelete_List_Click" Width="100px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDutyIndex_List" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDutyIndex_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyIndex") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepartment_List" runat="server" CssClass="text-Right-Blue" Text="部門別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="30%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eEmployee_List" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eEmpNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="30%" />
                            <asp:Label ID="eEmpName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDutyType_List" runat="server" CssClass="text-Right-Blue" Text="類別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDutyType_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyType_C") %>' Width="90%" />
                            <asp:Label ID="eDutyType_List" runat="server" Text='<%# Eval("DutyType") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:CheckBox ID="cbIsDisOneHour_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="扣除午休1小時" />
                            <asp:Label ID="eIsDisOneHour_List" runat="server" Text='<%# Eval("IsDisOneHour") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDutyDateStart_List" runat="server" CssClass="text-Right-Blue" Text="起迄日期時間：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eDutyDateStart_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyDateStart", "{0:yyyy/MM/dd}") %>' Width="25%" />
                            <asp:Label ID="eStartTime_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StartTime") %>' Width="10%" />
                            <asp:Label ID="lbSplit1_List" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                            <asp:Label ID="eDutyDateEnd_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyDateEnd", "{0:yyyy/MM/dd}") %>' Width="25%" />
                            <asp:Label ID="eEndTime_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EndTime") %>' Width="10%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDutyHours_List" runat="server" CssClass="text-Right-Blue" Text="時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="lbHours1_List" runat="server" CssClass="text-Right-Blue" Text="加班：" Width="10%" />
                            <asp:Label ID="eDutyHours_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DutyHours") %>' Width="30%" />
                            <asp:Label ID="lbHours2_List" runat="server" CssClass="text-Right-Blue" Text="補休：" Width="10%" />
                            <asp:Label ID="eESCHours_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ESCHours") %>' Width="30%" />
                            <asp:Label ID="eUsableHours_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UsableHours") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignReason_List" runat="server" CssClass="text-Right-Blue" Text="申請原由：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eAssignReason_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignReason") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbInspector_List" runat="server" CssClass="text-Right-Blue" Text="核准主管：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eInspector_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Inspector") %>' Width="30%" />
                            <asp:Label ID="eInspector_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Inspector_C") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="cbIsAllowed_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="已核准" />
                            <asp:Label ID="eIsAllowed_List" runat="server" Text='<%# Eval("IsAllowed") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="cbIsRejected_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="已駁回" />
                            <asp:Label ID="eIsRejected_List" runat="server" Text='<%# Eval("IsRejected") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColWidth-8Col">
                            <asp:Label ID="eHoursRatio_List" runat="server" Text='<%# Eval("HoursRatio") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRejectedReason_List" runat="server" CssClass="text-Right-Blue" Text="駁回原因：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRejectReason_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("RejectReason") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                            <asp:Label ID="eBuMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
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
        <asp:GridView ID="gridEmpDutyDataListB" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#DEDFDE" BorderStyle="None" BorderWidth="1px"
            CellPadding="4" DataKeyNames="DutyIndexItems" DataSourceID="sdsEmpDutyDataDetailB" ForeColor="Black" GridLines="Vertical" Width="50%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:BoundField DataField="DutyIndex" HeaderText="DutyIndex" SortExpression="DutyIndex" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                <asp:BoundField DataField="DutyIndexItems" HeaderText="DutyIndexItems" ReadOnly="True" SortExpression="DutyIndexItems" Visible="False" />
                <asp:BoundField DataField="SourceIndex" HeaderText="換休單號" SortExpression="SourceIndex" />
                <asp:BoundField DataField="ESCHours" HeaderText="換休時數" SortExpression="ESCHours" />
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
    <asp:SqlDataSource ID="sdsEmpDataList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT a.DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DEPNO)) AS DepName, a.TITLE, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '員工資料        EMPLOYEE        TITLE') AND (CLASSNO = a.TITLE)) AS Title_C, a.EMPNO, a.NAME, SUM(ISNULL(b.UsableHours, 0)) AS UseableHours FROM EMPLOYEE AS a LEFT OUTER JOIN EmpDutyHours AS b ON b.EmpNo = a.EMPNO AND b.IsAllowed = 1 WHERE (a.LEAVEDAY IS NULL) AND (ISNULL(a.DEPNO, '00') &lt;&gt; '00') AND (a.EMPNO &lt;&gt; 'supervisor') AND (DATEDIFF(Month, b.DutyDate, GETDATE()) &lt; 12) GROUP BY a.DEPNO, a.TITLE, a.EMPNO, a.NAME ORDER BY a.DEPNO, a.TITLE, a.EMPNO"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsEmpDutyDataList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT DutyIndex, DepNo, DepName, EmpNo, EmpName, DutyType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.DutyType) AND (FKEY = CAST('加班補休資料' AS char(16)) + CAST('fmEmpDutyHours' AS char(16)) + 'DutyType')) AS DutyType_C, DutyDateStart, StartTime, DutyDateEnd, EndTime, DutyHours, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuMan_C, BuDate, AssignReason, Remark, Inspector, (SELECT NAME FROM EMPLOYEE AS Employee_1 WHERE (EMPNO = a.Inspector)) AS Inspector_C, IsAllowed, RejectReason, IsRejected, HoursRatio, ESCHours, UsableHours, IsDisOneHour FROM EmpDutyHours AS a WHERE (EmpNo = @EmpNo) ORDER BY DutyDateStart DESC">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridEmpDataList" Name="EmpNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsEmpDutyDataDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT DutyIndex, DepNo, DepName, EmpNo, EmpName, DutyType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = CAST('加班補休資料' AS char(16)) + CAST('fmEmpDutyHours' AS char(16)) + 'DutyType') AND (CLASSNO = a.DutyType)) AS DutyType_C, DutyDateStart, StartTime, DutyDateEnd, EndTime, DutyHours, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuMan_C, BuDate, AssignReason, Remark, Inspector, (SELECT NAME FROM EMPLOYEE AS Employee_1 WHERE (EMPNO = a.Inspector)) AS Inspector_C, IsAllowed, RejectReason, IsRejected, HoursRatio, ESCHours, UsableHours, IsDisOneHour FROM EmpDutyHours AS a WHERE (DutyIndex = @DutyIndex)"
        InsertCommand="INSERT INTO EmpDutyHours(DutyIndex, DepNo, DepName, EmpNo, EmpName, DutyType, DutyDateStart, StartTime, DutyDateEnd, EndTime, DutyHours, BuMan, BuDate, AssignReason, Remark, Inspector, IsAllowed, IsRejected, RejectReason, HoursRatio, ESCHours, UsableHours, IsDisOneHour) VALUES (@DutyIndex, @DepNo, @DepName, @EmpNo, @EmpName, @DutyType, @DutyDateStart, @StartTime, @DutyDateEnd, @EndTime, @DutyHours, @BuMan, @BuDate, @AssignReason, @Remark, @Inspector, @IsAllowed, @IsRejected, @RejectReason, @HoursRatio, @ESCHours, @UsableHours, @IsDisOneHour)"
        UpdateCommand="UPDATE EmpDutyHours SET DutyDateStart = @DutyDateStart, StartTime = @StartTime, DutyDateEnd = @DutyDateEnd, EndTime = @EndTime, DutyHours = @DutyHours, AssignReason = @AssignReason, Remark = @Remark, ESCHours = @ESCHours, UsableHours = @UsableHours, Inspector = @Inspector, IsAllowed = @IsAllowed, IsRejected = @IsRejected, IsDisOneHour = @IsDisOneHour WHERE (DutyIndex = @DutyIndex)"
        DeleteCommand="DELETE FROM EmpDutyHours WHERE (DutyIndex = @DutyIndex)"
        OnDeleted="sdsEmpDutyDataDetail_Deleted">
        <DeleteParameters>
            <asp:Parameter Name="DutyIndex" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="DutyIndex" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="DepName" />
            <asp:Parameter Name="EmpNo" />
            <asp:Parameter Name="EmpName" />
            <asp:Parameter Name="DutyType" />
            <asp:Parameter Name="DutyDateStart" />
            <asp:Parameter Name="StartTime" />
            <asp:Parameter Name="DutyDateEnd" />
            <asp:Parameter Name="EndTime" />
            <asp:Parameter Name="DutyHours" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="BuDate" />
            <asp:Parameter Name="AssignReason" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="Inspector" />
            <asp:Parameter Name="IsAllowed" />
            <asp:Parameter Name="IsRejected" />
            <asp:Parameter Name="RejectReason" />
            <asp:Parameter Name="HoursRatio" />
            <asp:Parameter Name="ESCHours" />
            <asp:Parameter Name="UsableHours" />
            <asp:Parameter Name="IsDisOneHour" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridEmpDutyDataList" Name="DutyIndex" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="DutyDateStart" />
            <asp:Parameter Name="StartTime" />
            <asp:Parameter Name="DutyDateEnd" />
            <asp:Parameter Name="EndTime" />
            <asp:Parameter Name="DutyHours" />
            <asp:Parameter Name="AssignReason" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="ESCHours" />
            <asp:Parameter Name="UsableHours" />
            <asp:Parameter Name="DutyIndex" />
            <asp:Parameter Name="Inspector" />
            <asp:Parameter Name="IsAllowed" />
            <asp:Parameter Name="IsRejected" />
            <asp:Parameter Name="IsDisOneHour" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsEmpDutyDataDetailB" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT DutyIndex, Items, DutyIndexItems, SourceIndex, ESCHours FROM EmpDutyHoursB WHERE (DutyIndex = @DutyIndex)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridEmpDutyDataList" Name="DutyIndex" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsDutyType_INS" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = CAST('加班補休資料' AS char(16)) + CAST('fmEmpDutyHours' AS char(16)) + 'DutyType') ORDER BY ClassNo"></asp:SqlDataSource>
</asp:Content>
