<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="TicketCheckP.aspx.cs" Inherits="TyBus_Intranet_Test_V3.TicketCheckP" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="TicketCheckPForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">查票工作報告</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="colh ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCaseDate_Search" runat="server" CssClass="text-Right-Blue" Text="查票日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eCaseDate_S_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="lbSplit1_Search" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eCaseDate_E_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbInspector_Search" runat="server" CssClass="text-Right-Blue" Text="稽查人員：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eInspector_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eInspector_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eInspectorName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCarID_Search" runat="server" CssClass="text-Right-Blue" Text="牌照號碼：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCarID_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDriver_Search" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDriver_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eDriverName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbLinesNo_Search" runat="server" CssClass="text-Right-Blue" Text="路線：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eLinesNo_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:FileUpload ID="fuExcel" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" OnClick="bbSearch_Click" Text="搜尋" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbImportExcel" runat="server" CssClass="button-Black" Text="EXCEL 匯入" Width="95%" OnClick="bbImportExcel_Click" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" OnClick="bbClose_Click" Text="結束" Width="95%" />
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
        <asp:Button ID="bbPrint" runat="server" CssClass="button-Red" OnClick="bbPrint_Click" Text="列印報告" Width="120px" />
        <asp:Button ID="bbExportExcel" runat="server" CssClass="button-Blue" OnClick="bbExportExcel_Click" Text="匯出 EXCEL" Width="120px" />
        <br />
        <asp:GridView ID="gridTicketCheckReportList" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataKeyNames="CaseNo" DataSourceID="sdsTicketCheckReportList" GridLines="None" PageSize="5" Width="100%"
            OnPageIndexChanging="gridTicketCheckReportList_PageIndexChanging">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="CaseNo" HeaderText="單號" ReadOnly="True" SortExpression="CaseNo" />
                <asp:BoundField DataField="LinesNo" HeaderText="路線" SortExpression="LinesNo" />
                <asp:BoundField DataField="DepName" HeaderText="單位" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="Driver" HeaderText="駕駛員代號" SortExpression="Driver" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員姓名" SortExpression="DriverName" />
                <asp:BoundField DataField="CaseDate" DataFormatString="{0:D}" HeaderText="查票日期" SortExpression="CaseDate" />
                <asp:BoundField DataField="InspectorName" HeaderText="稽查人員" SortExpression="InspectorName" />
                <asp:BoundField DataField="Car_ID" HeaderText="車號" SortExpression="Car_ID" />
                <asp:BoundField DataField="DeparturePosition" HeaderText="出發地點" SortExpression="DeparturePosition" />
                <asp:BoundField DataField="DepartureTime" HeaderText="出發時間" SortExpression="DepartureTime" />
                <asp:BoundField DataField="ArrivePosition" HeaderText="到達地點" SortExpression="ArrivePosition" />
                <asp:BoundField DataField="ArriveTime" HeaderText="到達時間" SortExpression="ArriveTime" />
                <asp:BoundField DataField="CheckNote" HeaderText="查票經過" SortExpression="CheckNote" />
                <asp:BoundField DataField="CaseNote" HeaderText="處理經過" SortExpression="CaseNote" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" />
                <asp:CheckBoxField DataField="IsPassed" HeaderText="已開勸導單" SortExpression="IsPassed" />
                <asp:BoundField DataField="RP_ReportNo" HeaderText="勸導單號" SortExpression="RP_ReportNo" />
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
        <asp:FormView ID="fvTicketCheckReportDetail" runat="server" DataKeyNames="CaseNo" DataSourceID="sdsTicketCheckReportDetail" Width="100%" OnDataBound="fvTicketCheckReportDetail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Update" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseNo_Edit" runat="server" CssClass="text-Right-Blue" Text="單號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDriver_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Edit_TextChanged" Text='<%# Bind("Driver") %>' Width="35%" />
                                    <asp:Label ID="eDriverName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DriverName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="所屬部門：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Text='<%# Bind("DepNo") %>' Width="35%" />
                                    <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="cbHasAssetNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbHasAssetNo_Edit_CheckedChanged" Text="有財產編號" Width="95%" />
                                    <asp:Label ID="eHasAssetNo_Edit" runat="server" Text='<%# Bind("HasAssetNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="cbHasSecurityLabel_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbHasSecurityLabel_Edit_CheckedChanged" Text="有防偽標籤" Width="95%" />
                                    <asp:Label ID="eHasSecurityLabel_Edit" runat="server" Text='<%# Bind("HasSecurityLabel") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbStationToilet_Edit" runat="server" CssClass="text-Right-Blue" Text="場站廁所：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlStationToilet_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlStationToilet_Edit_SelectedIndexChanged" Width="95%"
                                        DataSourceID="sdsStationToilet_Edit" DataTextField="ClassTxt" DataValueField="ClassNo">
                                    </asp:DropDownList>
                                    <asp:TextBox ID="eStationToilet_Edit" runat="server" Text='<%# Bind("StationToilet") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbStationLight_Edit" runat="server" CssClass="text-Right-Blue" Text="場站燈光：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlStationLight_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnSelectedIndexChanged="ddlStationLight_Edit_SelectedIndexChanged" Width="95%">
                                        <asp:ListItem Text="" Value="" Selected="True" />
                                        <asp:ListItem Text="正常" Value="N" />
                                        <asp:ListItem Text="異常" Value="X" />
                                    </asp:DropDownList>
                                    <asp:TextBox ID="eStationLight_Edit" runat="server" Text='<%# Bind("StationLight") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseDate_Edit" runat="server" CssClass="text-Right-Blue" Text="查票日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCaseDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNo_Edit" runat="server" CssClass="text-Right-Blue" Text="路線別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLinesNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("LinesNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarID_Edit" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Car_ID") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="cbIsPassed_Edit" runat="server" CssClass="text-Right-Blue" Text="已發勸導單單號：" Enabled="false" Width="95%" />
                                    <asp:Label ID="eIsPassed_Edit" runat="server" Text='<%# Eval("IsPassed") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRP_ReportNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RP_ReportNo") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDeparturePosition_Edit" runat="server" CssClass="text-Right-Blue" Text="出發地點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDeparturePosition_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DeparturePosition") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepartureTime_Edit" runat="server" CssClass="text-Right-Blue" Text="出發時間：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDepartureTime_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DepartureTime") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbArrivePosition_Edit" runat="server" CssClass="text-Right-Blue" Text="到達地點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eArrivePosition_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ArrivePosition") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbArriveTime_Edit" runat="server" CssClass="text-Right-Blue" Text="到達時間：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eArriveTime_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ArriveTime") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCheckNote_Edit" runat="server" CssClass="text-Right-Blue" Text="查票經過：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eCheckNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("CheckNote") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseNote_Edit" runat="server" CssClass="text-Right-Blue" Text="處理經過：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eCaseNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("CaseNote") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbInspector_Edit" runat="server" CssClass="text-Right-Blue" Text="稽查人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eInspector_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eInspector_Edit_TextChanged" Text='<%# Bind("Inspector") %>' Width="35%" />
                                    <asp:Label ID="eInspectorName_Edit" runat="server" CssClass="text-Left-Black" Text='<%#Eval("InspectorName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-8Col" colspan="3" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
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
                                    <asp:Label ID="lbCaseNo_INS" runat="server" CssClass="text-Right-Blue" Text="單號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_INS" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDriver_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_INS_TextChanged" Text='<%# Bind("Driver") %>' Width="35%" />
                                    <asp:Label ID="eDriverName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DriverName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="所屬部門：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Bind("DepNo") %>' Width="35%" />
                                    <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="cbHasAssetNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnCheckedChanged="cbHasAssetNo_INS_CheckedChanged" Text="有財產編號" Width="95%" />
                                    <asp:Label ID="eHasAssetNo_INS" runat="server" Text='<%# Bind("HasAssetNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="cbHasSecurityLabel_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnCheckedChanged="cbHasSecurityLabel_INS_CheckedChanged" Text="有防偽標籤" Width="95%" />
                                    <asp:Label ID="eHasSecurityLabel_INS" runat="server" Text='<%# Bind("HasSecurityLabel") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbStationToilet_INS" runat="server" CssClass="text-Right-Blue" Text="場站廁所：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlStationToilet_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlStationToilet_INS_SelectedIndexChanged" Width="95%"
                                        DataSourceID="sdsStationToilet_INS" DataTextField="ClassTxt" DataValueField="ClassNo">
                                    </asp:DropDownList>
                                    <asp:TextBox ID="eStationToilet_INS" runat="server" Text='<%# Bind("StationToilet") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbStationLight_INS" runat="server" CssClass="text-Right-Blue" Text="場站燈光：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlStationLight_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnSelectedIndexChanged="ddlStationLight_INS_SelectedIndexChanged" Width="95%">
                                        <asp:ListItem Text="" Value="" Selected="True" />
                                        <asp:ListItem Text="正常" Value="N" />
                                        <asp:ListItem Text="異常" Value="X" />
                                    </asp:DropDownList>
                                    <asp:TextBox ID="eStationLight_INS" runat="server" Text='<%# Bind("StationLight") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseDate_INS" runat="server" CssClass="text-Right-Blue" Text="查票日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCaseDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNo_INS" runat="server" CssClass="text-Right-Blue" Text="路線別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLinesNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("LinesNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarID_INS" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Car_ID") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="cbIsPassed_INS" runat="server" CssClass="text-Right-Blue" Text="已發勸導單單號：" Enabled="false" Width="95%" />
                                    <asp:Label ID="eIsPassed_INS" runat="server" Text='<%# Eval("IsPassed") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRP_ReportNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RP_ReportNo") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDeparturePosition_INS" runat="server" CssClass="text-Right-Blue" Text="出發地點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDeparturePosition_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DeparturePosition") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepartureTime_INS" runat="server" CssClass="text-Right-Blue" Text="出發時間：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDepartureTime_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DepartureTime") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbArrivePosition_INS" runat="server" CssClass="text-Right-Blue" Text="到達地點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eArrivePosition_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ArrivePosition") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbArriveTime_INS" runat="server" CssClass="text-Right-Blue" Text="到達時間：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eArriveTime_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ArriveTime") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCheckNote_INS" runat="server" CssClass="text-Right-Blue" Text="查票經過：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eCheckNote_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("CheckNote") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseNote_INS" runat="server" CssClass="text-Right-Blue" Text="處理經過：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eCaseNote_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("CaseNote") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbInspector_INS" runat="server" CssClass="text-Right-Blue" Text="稽查人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eInspector_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eInspector_INS_TextChanged" Text='<%# Bind("Inspector") %>' Width="35%" />
                                    <asp:Label ID="eInspectorName_INS" runat="server" CssClass="text-Left-Black" Text='<%#Eval("InspectorName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-8Col" colspan="3" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
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
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="Edit" Text="更正" Width="120px" />
                &nbsp;<asp:Button ID="bbDel_List" runat="server" CausesValidation="False" CssClass="button-Red" OnClick="bbDel_List_Click" Text="刪除" Width="120px" />
                &nbsp;<asp:Button ID="bbCreateRPReport_List" runat="server" CausesValidation="false" CssClass="button-Red" OnClick="bbCreateRPReport_List_Click" Text="開立勸導單" Width="120px" />
                &nbsp;<asp:Button ID="bbPrintRPReport_List" runat="server" CausesValidation="false" CssClass="button-Black" OnClick="bbPrintRPReport_List_Click" Text="列印勸導單" Width="120px" />
                &nbsp;<asp:Button ID="bbPrintViolationReport_List" runat="server" CausesValidation="false" CssClass="button-Red" OnClick="bbPrintViolationReport_List_Click" Text="列印違規通知單" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseNo_List" runat="server" CssClass="text-Right-Blue" Text="單號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriver_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDriver_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver") %>' Width="35%" />
                            <asp:Label ID="eDriverName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="所屬部門：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:CheckBox ID="cbHasAssetNo_List" runat="server" CssClass="text-Left-Black" Text="有財產編號" Width="95%" />
                            <asp:Label ID="eHasAssetNo_List" runat="server" Text='<%# Eval("HasAssetNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:CheckBox ID="cbHasSecurityLabel_List" runat="server" CssClass="text-Left-Black" Text="有防偽標籤" Width="95%" />
                            <asp:Label ID="eHasSecurityLabel_List" runat="server" Text='<%# Eval("HasSecurityLabel") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStationToilet_List" runat="server" CssClass="text-Right-Blue" Text="場站廁所：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStationToilet_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StationToilet_C") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStationLight_List" runat="server" CssClass="text-Right-Blue" Text="場站燈光：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStationLight_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StationLight_C") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseDate_List" runat="server" CssClass="text-Right-Blue" Text="查票日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLinesNo_List" runat="server" CssClass="text-Right-Blue" Text="路線別：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLinesNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCarID_List" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCar_ID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="cbIsPassed_List" runat="server" CssClass="text-Right-Blue" Text="已發勸導單單號：" Enabled="false" Width="95%" />
                            <asp:Label ID="eIsPassed_List" runat="server" Text='<%# Eval("IsPassed") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eRP_ReportNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RP_ReportNo") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDeparturePosition_List" runat="server" CssClass="text-Right-Blue" Text="出發地點：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDeparturePosition_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DeparturePosition") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepartureTime_List" runat="server" CssClass="text-Right-Blue" Text="出發時間：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDepartureTime_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepartureTime") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbArrivePosition_List" runat="server" CssClass="text-Right-Blue" Text="到達地點：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eArrivePosition_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ArrivePosition") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbArriveTime_List" runat="server" CssClass="text-Right-Blue" Text="到達時間：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eArriveTime_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ArriveTime") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCheckNote_List" runat="server" CssClass="text-Right-Blue" Text="查票經過：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eCheckNote_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("CheckNote") %>' Height="95%" Width="98%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseNote_List" runat="server" CssClass="text-Right-Blue" Text="處理經過：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eCaseNote_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("CaseNote") %>' Height="95%" Width="98%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Height="95%" Width="98%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbInspector_List" runat="server" CssClass="text-Right-Blue" Text="稽查人員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eInspector_List" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Inspector") %>' Width="35%" />
                            <asp:Label ID="eInspectorName_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("InspectorName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-8Col" colspan="3" />
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
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
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
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
            <LocalReport ReportPath="Report\TicketCheckReportP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsTicketCheckReportList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, LinesNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = t.DepNo)) AS DepName, Driver, DriverName, CaseDate, DeparturePosition, DepartureTime, ArrivePosition, ArriveTime, Car_ID, CheckNote, CaseNote, Remark, IsPassed, RP_ReportNo, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = t.Inspector)) AS InspectorName, HasAssetNo, HasSecurityLabel, StationToilet, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '查票工作報告    TicketCheckP    StationToilet') AND (CLASSNO = t.StationToilet)) AS StationToilet_C, StationLight, CASE WHEN t.StationLight = 'X' THEN '異常' ELSE '正常' END AS StationLight_C FROM TicketCheckReport AS t WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsTicketCheckReportDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        DeleteCommand="DELETE FROM TicketCheckReport WHERE (CaseNo = @CaseNo)" InsertCommand="INSERT INTO TicketCheckReport(CaseNo, LinesNo, DepNo, Driver, DriverName, CaseDate, DeparturePosition, DepartureTime, ArrivePosition, ArriveTime, Car_ID, CheckNote, CaseNote, Remark, BuDate, BuMan, Inspector, HasAssetNo, HasSecurityLabel, StationToilet, StationLight) VALUES (@CaseNo, @LinesNo, @DepNo, @Driver, @DriverName, @CaseDate, @DeparturePosition, @DepartureTime, @ArrivePosition, @ArriveTime, @Car_ID, @CheckNote, @CaseNote, @Remark, @BuDate, @BuMan, @Inspector, @HasAssetNo, @HasSecurityLabel, @StationToilet, @StationLight)"
        SelectCommand="SELECT CaseNo, LinesNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = t.DepNo)) AS DepName, Driver, DriverName, CaseDate, DeparturePosition, DepartureTime, ArrivePosition, ArriveTime, Car_ID, CheckNote, CaseNote, Remark, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = t.BuMan)) AS BuManName, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = t.ModifyMan)) AS ModifyManName, IsPassed, RP_ReportNo, Inspector, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = t.Inspector)) AS InspectorName, HasAssetNo, HasSecurityLabel, StationToilet, CASE WHEN isnull(StationToilet , '') &lt;&gt; '' THEN (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '查票工作報告    TicketCheckP    StationToilet') AND (CLASSNO = t.StationToilet)) ELSE '正常' END AS StationToilet_C, StationLight, CASE WHEN t.StationLight = 'X' THEN '異常' ELSE '正常' END AS StationLight_C FROM TicketCheckReport AS t WHERE (CaseNo = @CaseNo)"
        UpdateCommand="UPDATE TicketCheckReport SET LinesNo = @LinesNo, DepNo = @DepNo, Driver = @Driver, DriverName = @DriverName, CaseDate = @CaseDate, DeparturePosition = @DeparturePosition, DepartureTime = @DepartureTime, ArrivePosition = @ArrivePosition, ArriveTime = @ArriveTime, Car_ID = @Car_ID, CheckNote = @CheckNote, CaseNote = @CaseNote, Remark = @Remark, ModifyDate = @ModifyDate, ModifyMan = @ModifyMan, Inspector = @Inspector, HasAssetNo = @HasAssetNo, HasSecurityLabel = @HasSecurityLabel, StationToilet = @StationToilet, StationLight = @StationLight WHERE (CaseNo = @CaseNo)"
        OnInserting="sdsTicketCheckReportDetail_Inserting"
        OnDeleted="sdsTicketCheckReportDetail_Deleted"
        OnInserted="sdsTicketCheckReportDetail_Inserted"
        OnUpdated="sdsTicketCheckReportDetail_Updated"
        OnUpdating="sdsTicketCheckReportDetail_Updating">
        <DeleteParameters>
            <asp:Parameter Name="CaseNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="LinesNo" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="DriverName" />
            <asp:Parameter Name="CaseDate" />
            <asp:Parameter Name="DeparturePosition" />
            <asp:Parameter Name="DepartureTime" />
            <asp:Parameter Name="ArrivePosition" />
            <asp:Parameter Name="ArriveTime" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="CheckNote" />
            <asp:Parameter Name="CaseNote" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuDate" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="Inspector" />
            <asp:Parameter Name="HasAssetNo" />
            <asp:Parameter Name="HasSecurityLabel" />
            <asp:Parameter Name="StationToilet" />
            <asp:Parameter Name="StationLight" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridTicketCheckReportList" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="LinesNo" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="DriverName" />
            <asp:Parameter Name="CaseDate" />
            <asp:Parameter Name="DeparturePosition" />
            <asp:Parameter Name="DepartureTime" />
            <asp:Parameter Name="ArrivePosition" />
            <asp:Parameter Name="ArriveTime" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="CheckNote" />
            <asp:Parameter Name="CaseNote" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="Inspector" />
            <asp:Parameter Name="HasAssetNo" />
            <asp:Parameter Name="HasSecurityLabel" />
            <asp:Parameter Name="StationToilet" />
            <asp:Parameter Name="StationLight" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsStationToilet_INS" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar) AS ClassTxt, CAST(NULL AS varchar) AS ClassNo UNION ALL SELECT CLASSTXT, CLASSNO FROM DBDICB WHERE (FKEY = '查票工作報告    TicketCheckP    StationToilet') ORDER BY CLassNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsStationToilet_Edit" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar) AS ClassTxt, CAST(NULL AS varchar) AS ClassNo UNION ALL SELECT CLASSTXT, CLASSNO FROM DBDICB WHERE (FKEY = '查票工作報告    TicketCheckP    StationToilet') ORDER BY CLassNo"></asp:SqlDataSource>
</asp:Content>
