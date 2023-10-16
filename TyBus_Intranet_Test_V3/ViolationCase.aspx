<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ViolationCase.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ViolationCase" %>

<asp:Content ID="ViolationCaseForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">違規案件處理</a>
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
                <td class="ColHeight ColWidth-8Col" colspan="2">
                    <asp:FileUpload ID="fuExcel" runat="server" Width="45%" />
                    <asp:Button ID="bbImportData" runat="server" CssClass="button-Red" Text="匯入資料" OnClick="bbImportData_Click" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出EXCEL" OnClick="bbExcel_Click" Width="95%" />
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
    <asp:Panel ID="plEditData" runat="server" CssClass="ShowPanel-Detail">
        <asp:FormView ID="fvViolationCaseMain" runat="server" DataKeyNames="CaseNo" DataSourceID="sdsViolationCase_Main" Width="100%"
            OnDataBound="fvViolationCaseMain_DataBound">
            <EditItemTemplate>
                <asp:Button ID="UpdateButton" runat="server" CssClass="button-Black" CausesValidation="True" CommandName="Update" Text="更新" Width="120px" />
                &nbsp;<asp:Button ID="UpdateCancelButton" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="titleText-Blue" colspan="8">
                                    <a>違　　規　　單</a>
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseNo_Edit" runat="server" CssClass="text-Right-Blue" Text="違規序號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseType_Edit" runat="server" CssClass="text-Right-Blue" Text="違規類別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCaseType_Edit" runat="server" Text='<%# Eval("CaseType") %>' OnTextChanged="eCaseType_Edit_TextChanged" Visible="false" />
                                    <asp:Label ID="eCaseTypeName_Edit" runat="server" CssClass="text-Left-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("BuildMan") %>' Width="35%" />
                                    <asp:Label ID="eBuildManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("BuildManName") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNo_Edit" runat="server" CssClass="text-Right-Blue" Text="路線代號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLinesNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("LinesNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DepNo") %>'
                                        AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Width="30%" />
                                    <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DepName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCar_ID_Edit" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Car_ID") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_Edit" runat="server" CssClass="text-Right-Blue" Text="違規駕駛：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDriver_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Driver") %>'
                                        AutoPostBack="true" OnTextChanged="eDriver_Edit_TextChanged" Width="30%" />
                                    <asp:Label ID="eDriverName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DriverName") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="titleText-S-Blue" colspan="8">
                                    <a>違規內容</a>
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbViolationDate_Edit" runat="server" CssClass="text-Right-Blue" Text="違規日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eViolationDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ViolationDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTicketNo_Edit" runat="server" CssClass="text-Right-Blue" Text="罰單號碼：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTicketNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("TicketNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPenaltyDep_Edit" runat="server" CssClass="text-Right-Blue" Text="裁罰單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="ePenaltyDep_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("PenaltyDep") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbUndertaker_Edit" runat="server" CssClass="text-Right-Blue" Text="承辦人(開單人)：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eUndertaker_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Undertaker") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbViolationRule_Edit" runat="server" CssClass="text-Right-Blue" Text="違規法條：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eViolationRule_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ViolationRule") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFineAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="裁罰金額：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eFineAmount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("FineAmount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPenaltyReason_Edit" runat="server" CssClass="text-Right-Blue" Text="裁罰原因：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="ePenaltyReason_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("PenaltyReason") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColBorder ColWidth-8Col MultiLine-High">
                                    <asp:Label ID="lbTicketTitle_Edit" runat="server" CssClass="text-Right-Blue" Text="裁罰主旨：" Width="95%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-High" colspan="7">
                                    <asp:TextBox ID="eTicketTitle_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("TicketTitle") %>'
                                        Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbViolationLocation_Edit" runat="server" CssClass="text-Right-Blue" Text="違規地點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eViolationLocation_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ViolationLocation") %>' Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-High ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbViolationNote_Edit" runat="server" CssClass="text-Right-Blue" Text="違規事項 (說明)：" Width="95%" />
                                </td>
                                <td class="MultiLine-High ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eViolationNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("ViolationNote") %>'
                                        Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbViolationPoint_Edit" runat="server" CssClass="text-Right-Blue" Text="記 (扣) 點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eViolationPoint_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ViolationPoint") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPaymentDeadline_Edit" runat="server" CssClass="text-Right-Blue" Text="繳納期限：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="ePaymentDeadline_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PaymentDeadline","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPaidDate_Edit" runat="server" CssClass="text-Right-Blue" Text="繳納日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="ePaidDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PaidDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColWidth-8Col" colspan="2" />
                            </tr>
                            <tr>
                                <td class="MultiLine-High ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine-High ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbInsert_Empty" runat="server" CssClass="button-Black" CausesValidation="true" CommandName="New" Text="新增" Width="120px" />
            </EmptyDataTemplate>
            <InsertItemTemplate>
                <asp:Button ID="InsertButton" runat="server" CssClass="button-Black" CausesValidation="True" CommandName="Insert" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="InsertCancelButton" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="titleText-Blue" colspan="8">
                                    <a>違　　規　　單</a>
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseNo_Insert" runat="server" CssClass="text-Right-Blue" Text="違規序號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseType_Insert" runat="server" CssClass="text-Right-Blue" Text="違規類別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlCaseType_Insert" runat="server" AutoPostBack="True"
                                        DataSourceID="sdsCaseType" DataTextField="ClassTxt" DataValueField="ClassNo"
                                        OnSelectedIndexChanged="ddlCaseType_Insert_SelectedIndexChanged" />
                                    <asp:TextBox ID="eCaseType_Insert" runat="server" Text='<%# Bind("CaseType") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildDate_Insert" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBuildDate_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildMan_Insert" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildMan_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("BuildMan") %>' Width="35%" />
                                    <asp:Label ID="eBuildManName_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("BuildManName") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNo_Insert" runat="server" CssClass="text-Right-Blue" Text="路線代號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLinesNo_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("LinesNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Insert" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDepNo_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DepNo") %>'
                                        AutoPostBack="true" OnTextChanged="eDepNo_Insert_TextChanged" Width="30%" />
                                    <asp:Label ID="eDepName_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DepName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCar_ID_Insert" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Car_ID") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_Insert" runat="server" CssClass="text-Right-Blue" Text="違規駕駛：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDriver_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Driver") %>'
                                        AutoPostBack="true" OnTextChanged="eDriver_Insert_TextChanged" Width="30%" />
                                    <asp:Label ID="eDriverName_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DriverName") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="titleText-S-Blue" colspan="8">
                                    <a>違規內容</a>
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbViolationDate_Insert" runat="server" CssClass="text-Right-Blue" Text="違規日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eViolationDate_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ViolationDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTicketNo_Insert" runat="server" CssClass="text-Right-Blue" Text="罰單號碼：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTicketNo_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("TicketNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPenaltyDep_Insert" runat="server" CssClass="text-Right-Blue" Text="裁罰單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="ePenaltyDep_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("PenaltyDep") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbUndertaker_Insert" runat="server" CssClass="text-Right-Blue" Text="承辦人(開單人)：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eUndertaker_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Undertaker") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbViolationRule_Insert" runat="server" CssClass="text-Right-Blue" Text="違規法條：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eViolationRule_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ViolationRule") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFineAmount_Insert" runat="server" CssClass="text-Right-Blue" Text="裁罰金額：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eFineAmount_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("FineAmount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPenaltyReason_Insert" runat="server" CssClass="text-Right-Blue" Text="裁罰原因：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="ePenaltyReason_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("PenaltyReason") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColBorder ColWidth-8Col MultiLine-High">
                                    <asp:Label ID="lbTicketTitle_Insert" runat="server" CssClass="text-Right-Blue" Text="裁罰主旨：" Width="95%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine-High" colspan="7">
                                    <asp:TextBox ID="eTicketTitle_Insert" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("TicketTitle") %>'
                                        Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbViolationLocation_Insert" runat="server" CssClass="text-Right-Blue" Text="違規地點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eViolationLocation_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ViolationLocation") %>' Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-High ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbViolationNote_Insert" runat="server" CssClass="text-Right-Blue" Text="違規事項 (說明)：" Width="95%" />
                                </td>
                                <td class="MultiLine-High ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eViolationNote_Insert" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("ViolationNote") %>'
                                        Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbViolationPoint_Insert" runat="server" CssClass="text-Right-Blue" Text="記 (扣) 點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eViolationPoint_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ViolationPoint") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPaymentDeadline_Insert" runat="server" CssClass="text-Right-Blue" Text="繳納期限：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="ePaymentDeadline_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PaymentDeadline","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPaidDate_Insert" runat="server" CssClass="text-Right-Blue" Text="繳納日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="ePaidDate_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PaidDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColWidth-8Col" colspan="2" />
                            </tr>
                            <tr>
                                <td class="MultiLine-High ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_Insert" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine-High ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_Insert" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>'
                                        Width="97%" Height="95%" />
                                </td>
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </InsertItemTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_List" Text="新增" runat="server" CssClass="button-Black" CausesValidation="false" CommandName="New" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit" Text="修改" runat="server" CssClass="button-Red" CausesValidation="false" CommandName="Edit" Width="120px" />
                &nbsp;<asp:Button ID="bbDelete" Text="刪除" runat="server" CssClass="button-Red" CausesValidation="false" OnClick="bbDelete_Click" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="titleText-Blue" colspan="8">
                            <a>違　　規　　單</a>
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseNo_List" runat="server" CssClass="text-Right-Blue" Text="違規序號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseType_List" runat="server" CssClass="text-Right-Blue" Text="違規類別：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseType_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseType_C") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuildDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuildMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="35%" />
                            <asp:Label ID="eBuildManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLinesNo_List" runat="server" CssClass="text-Right-Blue" Text="路線代號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLinesNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCar_ID_List" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCar_ID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriver_List" runat="server" CssClass="text-Right-Blue" Text="違規駕駛：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDriver_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver") %>' Width="35%" />
                            <asp:Label ID="eDriverName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="titleText-S-Blue" colspan="8">
                            <a>違規內容</a>
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbViolationDate_List" runat="server" CssClass="text-Right-Blue" Text="違規日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eViolationDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ViolationDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbTicketNo_List" runat="server" CssClass="text-Right-Blue" Text="罰單號碼：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eTicketNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TicketNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPenaltyDep_List" runat="server" CssClass="text-Right-Blue" Text="裁罰單位：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="ePenaltyDep_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PenaltyDep") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbUndertaker_List" runat="server" CssClass="text-Right-Blue" Text="承辦人(開單人)：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eUndertaker_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Undertaker") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbViolationRule_List" runat="server" CssClass="text-Right-Blue" Text="違規法條：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eViolationRule_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ViolationRule") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbFineAmount_List" runat="server" CssClass="text-Right-Blue" Text="裁罰金額：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eFineAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FineAmount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPenaltyReason_List" runat="server" CssClass="text-Right-Blue" Text="裁罰原因：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="ePenaltyReason_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PenaltyReason") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColBorder ColWidth-8Col MultiLine-High">
                            <asp:Label ID="lbTicketTitle_List" runat="server" CssClass="text-Right-Blue" Text="裁罰主旨：" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine-High" colspan="7">
                            <asp:TextBox ID="eTicketTitle_List" runat="server" CssClass="text-Left-Black" ReadOnly="true" TextMode="MultiLine" Text='<%# Eval("TicketTitle") %>'
                                Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbViolationLocation_List" runat="server" CssClass="text-Right-Blue" Text="違規地點：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="7">
                            <asp:Label ID="eViolationLocationLabel" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ViolationLocation") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-High ColBorder ColWidth-8Col">
                            <asp:Label ID="lbViolationNote_List" runat="server" CssClass="text-Right-Blue" Text="違規事項 (說明)：" Width="95%" />
                        </td>
                        <td class="MultiLine-High ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eViolationNote_List" runat="server" CssClass="text-Left-Black" ReadOnly="true" TextMode="MultiLine" Text='<%# Eval("ViolationNote") %>'
                                Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbViolationPoint_List" runat="server" CssClass="text-Right-Blue" Text="記 (扣) 點：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eViolationPoint_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ViolationPoint") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPaymentDeadline_List" runat="server" CssClass="text-Right-Blue" Text="繳納期限：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="ePaymentDeadline_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PaymentDeadline","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPaidDate_List" runat="server" CssClass="text-Right-Blue" Text="繳納日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="ePaidDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PaidDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColWidth-8Col" colspan="2" />
                    </tr>
                    <tr>
                        <td class="MultiLine-High ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                        </td>
                        <td class="MultiLine-High ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" ReadOnly="true" TextMode="MultiLine" Text='<%# Eval("Remark") %>'
                                Width="97%" Height="95%" />
                        </td>
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <br />
    <asp:Panel ID="plContactHis" runat="server" GroupingText="處理歷程明細" CssClass="ShowPanel">
        <asp:GridView ID="gridViolationCaseDetailList" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4"
            DataKeyNames="CaseNoItem" DataSourceID="sdsViolationCaseDetailList" ForeColor="#333333" GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
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
    </asp:Panel>
    <asp:SqlDataSource ID="sdsViolationCase_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, CaseType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = ViolationCase.CaseType) AND (FKEY = '違規記錄檔      ViolationCase   CaseType')) AS CaseType_C, BuildDate, BuildManName, LinesNo, DepName, Car_ID, DriverName, TicketTitle, PenaltyDep, FineAmount, ViolationDate FROM ViolationCase WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsViolationCase_Main" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, CaseType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = CAST('違規記錄檔' AS char(16)) + CAST('ViolationCase' AS char(16)) + 'CaseType') AND (CLASSNO = a.CaseType)) AS CaseType_C, BuildDate, BuildMan, BuildManName, LinesNo, DepNo, DepName, Car_ID, Driver, DriverName, ViolationDate, TicketNo, TicketTitle, PenaltyDep, Undertaker, ViolationLocation, ViolationRule, ViolationNote, FineAmount, PenaltyReason, ViolationPoint, PaymentDeadline, PaidDate, Remark FROM ViolationCase AS a WHERE (CaseNo = @CaseNo)"
        DeleteCommand="DELETE FROM ViolationCase WHERE (CaseNo = @CaseNo)"
        InsertCommand="INSERT INTO ViolationCase(CaseNo, CaseType, BuildDate, BuildMan, BuildManName, LinesNo, DepNo, DepName, Car_ID, Driver, DriverName, ViolationDate, TicketNo, TicketTitle, PenaltyDep, Undertaker, ViolationLocation, ViolationRule, ViolationNote, FineAmount, PenaltyReason, ViolationPoint, PaymentDeadline, PaidDate, Remark) VALUES (@CaseNo, @CaseType, @BuildDate, @BuildMan, @BuildManName, @LinesNo, @DepNo, @DepName, @Car_ID, @Driver, @DriverName, @ViolationDate, @TicketNo, @TicketTitle, @PenaltyDep, @Undertaker, @ViolationLocation, @ViolationRule, @ViolationNote, @FineAmount, @PenaltyReason, @ViolationPoint, @PaymentDeadline, @PaidDate, @Remark)"
        UpdateCommand="UPDATE ViolationCase SET LinesNo = @LinesNo, DepNo = @DepNo, DepName = @DepName, Car_ID = @Car_ID, Driver = @Driver, DriverName = @DriverName, ViolationDate = @ViolationDate, TicketNo = @TicketNo, TicketTitle = @TicketTitle, PenaltyDep = @PenaltyDep, Undertaker = @Undertaker, ViolationLocation = @ViolationLocation, ViolationRule = @ViolationRule, ViolationNote = @ViolationNote, FineAmount = @FineAmount, PenaltyReason = @PenaltyReason, ViolationPoint = @ViolationPoint, PaymentDeadline = @PaymentDeadline, PaidDate = @PaidDate, Remark = @Remark WHERE (CaseNo = @CaseNo)"
        OnInserting="sdsViolationCase_Main_Inserting"
        OnUpdating="sdsViolationCase_Main_Updating">
        <DeleteParameters>
            <asp:Parameter Name="CaseNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="CaseType" />
            <asp:Parameter Name="BuildDate" />
            <asp:Parameter Name="BuildMan" />
            <asp:Parameter Name="BuildManName" />
            <asp:Parameter Name="LinesNo" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="DepName" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="DriverName" />
            <asp:Parameter Name="ViolationDate" />
            <asp:Parameter Name="TicketNo" />
            <asp:Parameter Name="TicketTitle" />
            <asp:Parameter Name="PenaltyDep" />
            <asp:Parameter Name="Undertaker" />
            <asp:Parameter Name="ViolationLocation" />
            <asp:Parameter Name="ViolationRule" />
            <asp:Parameter Name="ViolationNote" />
            <asp:Parameter Name="FineAmount" />
            <asp:Parameter Name="PenaltyReason" />
            <asp:Parameter Name="ViolationPoint" />
            <asp:Parameter Name="PaymentDeadline" />
            <asp:Parameter Name="PaidDate" />
            <asp:Parameter Name="Remark" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridViolationCaseList" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="LinesNo" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="DepName" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="DriverName" />
            <asp:Parameter Name="ViolationDate" />
            <asp:Parameter Name="TicketNo" />
            <asp:Parameter Name="TicketTitle" />
            <asp:Parameter Name="PenaltyDep" />
            <asp:Parameter Name="Undertaker" />
            <asp:Parameter Name="ViolationLocation" />
            <asp:Parameter Name="ViolationRule" />
            <asp:Parameter Name="ViolationNote" />
            <asp:Parameter Name="FineAmount" />
            <asp:Parameter Name="PenaltyReason" />
            <asp:Parameter Name="ViolationPoint" />
            <asp:Parameter Name="PaymentDeadline" />
            <asp:Parameter Name="PaidDate" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="CaseNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsViolationCaseDetailList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNoItem, CaseNo, Items, ContactPerson, ContactNote, ContactDate, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = B.Excutort)) AS Excutort_C, AssignDate, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (EMPNO = B.AssignedMan)) AS AssignedMan_C FROM ViolationCaseB AS B WHERE (CaseNo = @CaseNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridViolationCaseList" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsCaseType" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar) AS ClassNo, CAST(NULL AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = CAST('違規記錄檔' AS char(16)) + CAST('ViolationCase' AS char(16)) + 'CaseType') ORDER BY ClassNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsCaseType_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar) AS ClassNo, CAST(NULL AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = CAST('違規記錄檔' AS char(16)) + CAST('ViolationCase' AS char(16)) + 'CaseType') ORDER BY ClassNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsImportDataFromExcel" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        InsertCommand="INSERT INTO ViolationCase(CaseNo, CaseType, BuildDate, BuildMan, BuildManName, LinesNo, DepNo, DepName, Car_ID, Driver, DriverName, ViolationDate, TicketNo, TicketTitle, PenaltyDep, Undertaker, ViolationLocation, ViolationRule, ViolationNote, FineAmount, PenaltyReason, ViolationPoint, PaymentDeadline, PaidDate, Remark) VALUES (@CaseNo, @CaseType, @BuildDate, @BuildMan, @BuildManName, @LinesNo, @DepNo, @DepName, @Car_ID, @Driver, @DriverName, @ViolationDate, @TicketNo, @TicketTitle, @PenaltyDep, @Undertaker, @ViolationLocation, @ViolationRule, @ViolationNote, @FineAmount, @PenaltyReason, @ViolationPoint, @PaymentDeadline, @PaidDate, @Remark)"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>">
        <InsertParameters>
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="CaseType" />
            <asp:Parameter Name="BuildDate" />
            <asp:Parameter Name="BuildMan" />
            <asp:Parameter Name="BuildManName" />
            <asp:Parameter Name="LinesNo" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="DepName" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="DriverName" />
            <asp:Parameter Name="ViolationDate" />
            <asp:Parameter Name="TicketNo" />
            <asp:Parameter Name="TicketTitle" />
            <asp:Parameter Name="PenaltyDep" />
            <asp:Parameter Name="Undertaker" />
            <asp:Parameter Name="ViolationLocation" />
            <asp:Parameter Name="ViolationRule" />
            <asp:Parameter Name="ViolationNote" />
            <asp:Parameter Name="FineAmount" />
            <asp:Parameter Name="PenaltyReason" />
            <asp:Parameter Name="ViolationPoint" />
            <asp:Parameter Name="PaymentDeadline" />
            <asp:Parameter Name="PaidDate" />
            <asp:Parameter Name="Remark" />
        </InsertParameters>
    </asp:SqlDataSource>
</asp:Content>
