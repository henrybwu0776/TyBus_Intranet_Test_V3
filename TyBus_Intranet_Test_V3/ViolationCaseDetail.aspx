<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ViolationCaseDetail.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ViolationCaseDetail" %>

<asp:Content ID="ViolationCaseDetailForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">違規案件處理歷程</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCaseType_Search" runat="server" CssClass="text-Right-Blue" Text="違規類別：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlCaseType_Search" runat="server" CssClass="text-Left-Black" Width="95%" AutoPostBack="True"
                        DataSourceID="sdsCaseType_Search" DataTextField="ClassTxt" DataValueField="ClassNo"
                        OnSelectedIndexChanged="ddlCaseType_Search_SelectedIndexChanged" />
                    <br />
                    <asp:TextBox ID="eCaseType_Search" runat="server" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit1_Search" runat="server" CssClass="titleText-S-Blue" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCarID_Search" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eCarID_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit2_Search" runat="server" CssClass="titleText-S-Blue" Text="～" Width="5%" />
                    <asp:TextBox ID="eCarID_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbViolationDate_Search" runat="server" CssClass="text-Right-Blue" Text="違規 (發文) 日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eViolationDate_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit3_Search" runat="server" CssClass="titleText-S-Blue" Text="～" Width="5%" />
                    <asp:TextBox ID="eViolationDate_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
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
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridViolationCaseList" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White"
            BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="CaseNo" DataSourceID="sdsViolationCase_List"
            GridLines="None" PageSize="5" Width="100%">
            <Columns>
                <asp:CommandField ShowSelectButton="True" ButtonType="Button" />
                <asp:BoundField DataField="CaseNo" HeaderText="違規序號" ReadOnly="True" SortExpression="CaseNo" />
                <asp:BoundField DataField="CaseType" HeaderText="CaseType" SortExpression="CaseType" Visible="false" />
                <asp:BoundField DataField="CaseType_C" HeaderText="違規種類" ReadOnly="True" SortExpression="CaseType_C" />
                <asp:BoundField DataField="ViolationDate" HeaderText="違規日期" SortExpression="ViolationDate" DataFormatString="{0:d}" />
                <asp:BoundField DataField="BuildDate" HeaderText="建檔日期" SortExpression="BuildDate" DataFormatString="{0:d}" />
                <asp:BoundField DataField="BuildManName" HeaderText="建檔人" SortExpression="BuildManName" />
                <asp:BoundField DataField="LinesNo" HeaderText="路線" SortExpression="LinesNo" />
                <asp:BoundField DataField="DepName" HeaderText="站別" SortExpression="DepName" />
                <asp:BoundField DataField="Car_ID" HeaderText="車號" SortExpression="Car_ID" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員" SortExpression="DriverName" />
                <asp:BoundField DataField="TicketTitle" HeaderText="發文主旨" SortExpression="TicketTitle" />
                <asp:BoundField DataField="PenaltyDep" HeaderText="裁罰單位" SortExpression="PenaltyDep" />
                <asp:BoundField DataField="FineAmount" HeaderText="裁罰金額" SortExpression="FineAmount" />
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
    <asp:Panel ID="DetailDataShow" runat="server" CssClass="ShowPanel-Detail">
        <asp:GridView ID="gridViolationCaseDetailList" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4"
            DataKeyNames="CaseNoItem" DataSourceID="sdsViolationCaseDetailList" ForeColor="#333333" GridLines="None" Width="100%">
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
        <asp:FormView ID="fvViolationCaseDetail" runat="server" DataKeyNames="CaseNoItem" DataSourceID="sdsViolationCaseDetail" Width="100%"
            OnDataBound="fvViolationCaseDetail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Update" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseNo_Edit" runat="server" CssClass="text-Right-Blue" Text="違規單號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="95%" />
                                    <asp:Label ID="eCaseNoItem_Edit" runat="server" Text='<%# Eval("CaseNoItem") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbItems_Edit" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="95%" />
                                    <br />
                                    <asp:Label ID="eBuildMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan_C") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbExcutort_Edit" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eExcutort_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eExcutort_Edit_TextChanged" Text='<%# Bind("Excutort") %>' Width="35%" />
                                    <asp:Label ID="eExcutort_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Excutort_C") %>' Width="60%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" rowspan="3">
                                    <asp:Label ID="lbContactNote_Edit" runat="server" CssClass="text-Right-Blue" Text="處理概述：" Width="95%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="3" rowspan="3">
                                    <asp:TextBox ID="eContactNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("ContactNote") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbContextDate_Edit" runat="server" CssClass="text-Right-Blue" Text="連絡日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eContactDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContactDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbContactPerson_Edit" runat="server" CssClass="text-Right-Blue" Text="連絡對象：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eContactPerson_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ContactPerson") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignDate_Edit" runat="server" CssClass="text-Right-Blue" Text="轉交日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Assigndate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignedMan_Edit" runat="server" CssClass="text-Right-Blue" Text="轉交人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignedMan_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignedMan_Edit_TextChanged" Text='<%# Bind("AssignedMan") %>' Width="95%" />
                                    <br />
                                    <asp:Label ID="eAssignedMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignedMan_C") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="95%" />
                                    <br />
                                    <asp:Label ID="eModifyMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="95%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" rowspan="2">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="異動說明：" Width="95%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="5" rowspan="2">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Insert" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseNo_INS" runat="server" CssClass="text-Right-Blue" Text="違規單號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_INS" runat="server" CssClass="text-Left-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbItems_INS" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eItems_INS" runat="server" CssClass="text-Left-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildDate_INS" runat="server" CssClass="text-Left-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildMan_INS" runat="server" CssClass="text-Left-Black" Width="95%" />
                                    <br />
                                    <asp:Label ID="eBuildMan_C_INS" runat="server" CssClass="text-Left-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbExcutort_INS" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eExcutort_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eExcutort_INS_TextChanged" Text='<%# Bind("Excutort") %>' Width="35%" />
                                    <asp:Label ID="eExcutort_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Excutort_C") %>' Width="60%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" rowspan="3">
                                    <asp:Label ID="lbContactNote_INS" runat="server" CssClass="text-Right-Blue" Text="處理概述：" Width="95%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="3" rowspan="3">
                                    <asp:TextBox ID="eContactNote_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("ContactNote") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbContextDate_INS" runat="server" CssClass="text-Right-Blue" Text="連絡日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eContactDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContactDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbContactPerson_INS" runat="server" CssClass="text-Right-Blue" Text="連絡對象：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eContactPerson_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ContactPerson") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignDate_INS" runat="server" CssClass="text-Right-Blue" Text="轉交日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Assigndate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignedMan_INS" runat="server" CssClass="text-Right-Blue" Text="轉交人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignedMan_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignedMan_INS_TextChanged" Text='<%# Bind("AssignedMan") %>' Width="95%" />
                                    <br />
                                    <asp:Label ID="eAssignedMan_C_INS" runat="server" CssClass="text-Left-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="95%" />
                                    <br />
                                    <asp:Label ID="eModifyMan_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="95%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" rowspan="2">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="異動說明：" Width="95%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="5" rowspan="2">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Enabled="false" Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
                <asp:Button ID="bbCreate_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="Edit" Text="編輯" Width="120px" />
                &nbsp;<asp:Button ID="bbDelete_List" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Delete" Text="刪除" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseNo_List" runat="server" CssClass="text-Right-Blue" Text="違規單號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseNoItem_List" runat="server" Text='<%# Eval("CaseNoItem") %>' Visible="false" />
                            <asp:Label ID="eCaseNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItems_List" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuildDate_LIst" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuildMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="95%" />
                            <br />
                            <asp:Label ID="eBuildMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan_C") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbExcutort_List" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eExcutort_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Excutort") %>' Width="35%" />
                            <asp:Label ID="eExcutort_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Excutort_C") %>' Width="60%" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine-Low" rowspan="3">
                            <asp:Label ID="lbContactNote_List" runat="server" CssClass="text-Right-Blue" Text="處理概述：" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="3" rowspan="3">
                            <asp:TextBox ID="eContactNote_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("ContactNote") %>' Enabled="false" Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbContextDate_List" runat="server" CssClass="text-Right-Blue" Text="連絡日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eContactDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContactDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbContactPerson_List" runat="server" CssClass="text-Right-Blue" Text="連絡對象：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eContactPerson_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContactPerson") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignDate_List" runat="server" CssClass="text-Right-Blue" Text="轉交日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAssignDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Assigndate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignedMan_List" runat="server" CssClass="text-Right-Blue" Text="轉交人員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAssignedMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignedMan") %>' Width="95%" />
                            <br />
                            <asp:Label ID="eAssignedMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignedMan_C") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="95%" />
                            <br />
                            <asp:Label ID="eModifyMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine-Low" rowspan="2">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="異動說明：" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="5" rowspan="2">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Enabled="false" Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
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
            <EmptyDataTemplate>
                <asp:Button ID="bbCreate_Empty" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" />
            </EmptyDataTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsViolationCase_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, CaseType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = ViolationCase.CaseType) AND (FKEY = '違規記錄檔      ViolationCase   CaseType')) AS CaseType_C, BuildDate, BuildManName, LinesNo, DepName, Car_ID, DriverName, TicketTitle, PenaltyDep, FineAmount, ViolationDate FROM ViolationCase WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsViolationCaseDetailList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNoItem, CaseNo, Items, ContactPerson, ContactNote, ContactDate, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = B.Excutort)) AS Excutort_C, AssignDate, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (EMPNO = B.AssignedMan)) AS AssignedMan_C FROM ViolationCaseB AS B WHERE (CaseNo = @CaseNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridViolationCaseList" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsViolationCaseDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        DeleteCommand="DELETE FROM ViolationCaseB WHERE (CaseNoItem = @CaseNoItem)"
        InsertCommand="INSERT INTO ViolationCaseB(CaseNoItem, CaseNo, Items, ContactPerson, ContactDate, ContactNote, Excutort, AssignDate, AssignedMan, BuildDate, BuildMan) VALUES (@CaseNoItem, @CaseNo, @Items, @ContactPerson, @ContactDate, @ContactNote, @Excutort, @AssignDate, @AssignedMan, @BuildDate, @BuildMan)"
        SelectCommand="SELECT CaseNoItem, CaseNo, Items, ContactPerson, ContactNote, ContactDate, Excutort, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = B.Excutort)) AS Excutort_C, AssignDate, AssignedMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (EMPNO = B.AssignedMan)) AS AssignedMan_C, BuildDate, BuildMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = B.BuildMan)) AS BuildMan_C, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = B.ModifyMan)) AS ModifyMan_C, Remark FROM ViolationCaseB AS B WHERE (CaseNoItem = @CaseNoItem)"
        UpdateCommand="UPDATE ViolationCaseB SET ContactPerson = @ContactPerson, ContactNote = @ContactNote, ContactDate = @ContactDate, Excutort = @Excutort, AssignDate = @AssignDate, AssignedMan = @AssignedMan, ModifyDate = @ModifyDate, ModifyMan = @ModifyMan, Remark = @Remark WHERE (CaseNoItem = @CaseNoItem)"
        OnDeleted="sdsViolationCaseDetail_Deleted"
        OnInserted="sdsViolationCaseDetail_Inserted"
        OnInserting="sdsViolationCaseDetail_Inserting"
        OnUpdated="sdsViolationCaseDetail_Updated"
        OnUpdating="sdsViolationCaseDetail_Updating">
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
            <asp:ControlParameter ControlID="gridViolationCaseDetailList" Name="CaseNoItem" PropertyName="SelectedValue" />
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
    <asp:SqlDataSource ID="sdsCaseType_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar) AS ClassNo, CAST(NULL AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = CAST('違規記錄檔' AS char(16)) + CAST('ViolationCase' AS char(16)) + 'CaseType') ORDER BY ClassNo"></asp:SqlDataSource>
</asp:Content>
