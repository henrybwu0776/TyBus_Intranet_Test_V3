<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ViolationCaseHistory.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ViolationCaseHistory" %>

<asp:Content ID="ViolationCaseHistoryForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">違規案件異動記錄</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuildDate" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eBuildDate_Start" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eBuildDate_End" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbViolationDate" runat="server" CssClass="text-Right-Blue" Text="違規日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eViolationDate_Start" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit2" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eViolationDate_End" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColWidth-8Col"></td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="站別編號：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Start" runat="server" CssClass="text-Left-Black" OnTextChanged="eDepNo_Start_TextChanged" AutoPostBack="true" Width="40%" />
                    <asp:Label ID="lbSplit3" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_End" runat="server" CssClass="text-Left-Black" OnTextChanged="eDepNo_End_TextChanged" AutoPostBack="true" Width="40%" />
                    <br />
                    <asp:Label ID="eDepName_Start" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbBlank" runat="server" CssClass="text-Left-Black" Text="" Width="5%" />
                    <asp:Label ID="eDepName_End" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDriver" runat="server" CssClass="text-Right-Blue" Text="違規駕駛編號：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eDriver" runat="server" CssClass="text-Left-Black" OnTextChanged="eDriver_TextChanged" AutoPostBack="true" Width="95%" />
                    <br />
                    <asp:Label ID="eDriverName" runat="server" CssClass="text-Left-Black" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCar_ID" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCar_ID" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="95%" />
                </td>
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
        <asp:GridView ID="gridViolationCaseHistory" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1"
            DataKeyNames="HistoryNo" DataSourceID="sdsViolationCaseHistoryList" GridLines="None" PageSize="5" Width="100%">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="HistoryNo" HeaderText="異動編號" InsertVisible="False" ReadOnly="True" SortExpression="HistoryNo" Visible="False" />
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
                <asp:BoundField DataField="ModifyType" HeaderText="ModifyType" SortExpression="ModifyType" Visible="False" />
                <asp:BoundField DataField="ModifyType_C" HeaderText="異動類別" ReadOnly="True" SortExpression="ModifyType_C" />
                <asp:BoundField DataField="ModifyDate" DataFormatString="{0:d}" HeaderText="異動日期" SortExpression="ModifyDate" />
                <asp:BoundField DataField="ModifyMan" HeaderText="ModifyMan" SortExpression="ModifyMan" Visible="False" />
                <asp:BoundField DataField="ModifyMan_C" HeaderText="異動人員" ReadOnly="True" SortExpression="ModifyMan_C" />
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
        <asp:FormView ID="fvViolationCaseHistory" runat="server" DataKeyNames="HistoryNo" DataSourceID="sdsViolationCaseHistoryDetail" Width="100%">
            <ItemTemplate>
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
                            <asp:Label ID="lbLinesNo" runat="server" CssClass="text-Right-Blue" Text="路線代號：" Width="95%" />
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
                            <asp:Label ID="eHistoryNo_List" runat="server" Text='<%# Eval("HistoryNo") %>' Visible="false" />
                        </td>
                        <td class="MultiLine-High ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" ReadOnly="true" TextMode="MultiLine" Text='<%# Eval("Remark") %>'
                                Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyType_List" runat="server" CssClass="text-Right-Blue" Text="異動說明：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="7">
                            <asp:Label ID="eModifyType_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyType_C") %>' />
                            &nbsp;
                <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' />
                            &nbsp;
                <asp:Label ID="eModifyMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' />
                        </td>
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsViolationCaseHistoryList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT HistoryNo, CaseNo, CaseType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.CaseType) AND (FKEY = CAST('違規記錄檔' AS char(16)) + CAST('ViolationCase' AS char(16)) + 'CaseType')) AS CaseType_C, BuildDate, BuildMan, BuildManName, LinesNo, DepNo, DepName, Car_ID, Driver, DriverName, ViolationDate, TicketNo, TicketTitle, PenaltyDep, Undertaker, ViolationLocation, ViolationRule, ViolationNote, FineAmount, PenaltyReason, ViolationPoint, PaymentDeadline, PaidDate, Remark, ModifyType, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (CLASSNO = a.ModifyType) AND (FKEY = CAST('客訴單歷史' AS char(16)) + CAST('fmServiceHistory' AS char(16)) + 'ModifyType')) AS ModifyType_C, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.ModifyMan)) AS ModifyMan_C FROM ViolationCaseHistory AS a"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsViolationCaseHistoryDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT HistoryNo, CaseNo, CaseType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.CaseType) AND (FKEY = CAST('違規記錄檔' AS char(16)) + CAST('ViolationCase' AS char(16)) + 'CaseType')) AS CaseType_C, BuildDate, BuildMan, BuildManName, LinesNo, DepNo, DepName, Car_ID, Driver, DriverName, ViolationDate, TicketNo, TicketTitle, PenaltyDep, Undertaker, ViolationLocation, ViolationRule, ViolationNote, FineAmount, PenaltyReason, ViolationPoint, PaymentDeadline, PaidDate, Remark, ModifyType, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (CLASSNO = a.ModifyType) AND (FKEY = CAST('客訴單歷史' AS char(16)) + CAST('fmServiceHistory' AS char(16)) + 'ModifyType')) AS ModifyType_C, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.ModifyMan)) AS ModifyMan_C FROM ViolationCaseHistory AS a WHERE (HistoryNo = @HistoryNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridViolationCaseHistory" Name="HistoryNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
</asp:Content>
