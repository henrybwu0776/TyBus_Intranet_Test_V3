<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="RunsheetDailyCheck.aspx.cs" Inherits="TyBus_Intranet_Test_V3.RunsheetDailyCheck" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="RunsheetDailyCheckForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">每日憑單及碼表抽查表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuDate_Search" runat="server" CssClass="text-Right-Blue" Text="開單日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eBuDateS_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eBuDateE_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCaseDate_Search" runat="server" CssClass="text-Right-Blue" Text="憑單或碼表日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eCaseDateS_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eCaseDateE_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuMan_Search" runat="server" CssClass="text-Right-Blue" Text="開單人員：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eBuMan_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eBuMan_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eBuManName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCarID_Search" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCarID_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eDepName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDriver_Search" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDriver_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eDriverName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" OnClick="bbSearch_Click" Text="搜尋" Width="120px" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" OnClick="bbClose_Click" Text="結束" Width="120px" />
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
        <asp:Button ID="bbPrint" runat="server" CssClass="button-Red" OnClick="bbPrint_Click" Text="列印報表" Width="120px" />
        <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" OnClick="bbExcel_Click" Text="匯出 EXCEL" Width="120px" />
        <asp:GridView ID="gridRunsheetDailyCheck_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="CaseNo" DataSourceID="sdsRunsheetDailyCheck_List" GridLines="None" Width="100%">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="CaseNo" HeaderText="CaseNo" ReadOnly="True" SortExpression="CaseNo" Visible="False" />
                <asp:BoundField DataField="CaseDate" DataFormatString="{0:D}" HeaderText="憑單及碼表日期" SortExpression="CaseDate" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="Car_ID" HeaderText="車號" SortExpression="Car_ID" />
                <asp:BoundField DataField="Driver" HeaderText="Driver" SortExpression="Driver" Visible="False" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員" SortExpression="DriverName" />
                <asp:BoundField DataField="Remark" HeaderText="錯誤原因" SortExpression="Remark" />
                <asp:BoundField DataField="Inspector" HeaderText="Inspector" SortExpression="Inspector" Visible="False" />
                <asp:BoundField DataField="InspectorName" HeaderText="查核人" ReadOnly="True" SortExpression="InspectorName" />
                <asp:BoundField DataField="BuDate" DataFormatString="{0:D}" HeaderText="開單日期" SortExpression="BuDate" />
                <asp:BoundField DataField="BuMan" HeaderText="BuMan" SortExpression="BuMan" Visible="False" />
                <asp:BoundField DataField="BuManName" HeaderText="開單人" ReadOnly="True" SortExpression="BuManName" />
                <asp:BoundField DataField="ModifyDate" DataFormatString="{0:D}" HeaderText="異動日期" SortExpression="ModifyDate" />
                <asp:BoundField DataField="ModifyMan" HeaderText="ModifyMan" SortExpression="ModifyMan" Visible="False" />
                <asp:BoundField DataField="ModifyManName" HeaderText="異動人" ReadOnly="True" SortExpression="ModifyManName" />
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
        <asp:FormView ID="fvRunsheetDailyCheck_Detail" runat="server" DataKeyNames="CaseNo" DataSourceID="sdsRunsheetDailyCheck_Detail" Width="100%"
            OnDataBound="fvRunsheetDailyCheck_Detail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Update" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseDate_Edit" runat="server" CssClass="text-Right-Blue" Text="憑單及碼表日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCaseDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_Edit" runat="server" Text='<%# Eval("CaseNo") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Text='<%# Bind("DepNo") %>' Width="35%" />
                                    <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCar_ID_Edit" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Car_ID") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDriver_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Edit_TextChanged" Text='<%# Bind("Driver") %>' Width="35%" />
                                    <asp:Label ID="eDriverName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DriverName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="錯誤原因：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbInspector_Edit" runat="server" CssClass="text-Right-Blue" Text="查核人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eInspector_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eInspector_Edit_TextChanged" Text='<%# Bind("Inspector") %>' Width="35%" />
                                    <asp:Label ID="eInspectorName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InspectorName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColBorder">
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
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CommandName="Insert" CssClass="button-Black" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CommandName="Cancel" CssClass="button-Red" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseDate_INS" runat="server" CssClass="text-Right-Blue" Text="憑單及碼表日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCaseDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_INS" runat="server" Text='<%# Eval("CaseNo") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Bind("DepNo") %>' Width="35%" />
                                    <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCar_ID_INS" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Car_ID") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_INS" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDriver_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_INS_TextChanged" Text='<%# Bind("Driver") %>' Width="35%" />
                                    <asp:Label ID="eDriverName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DriverName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="錯誤原因：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbInspector_INS" runat="server" CssClass="text-Right-Blue" Text="查核人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eInspector_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eInspector_INS_TextChanged" Text='<%# Bind("Inspector") %>' Width="35%" />
                                    <asp:Label ID="eInspectorName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InspectorName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColBorder">
                                    <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ModifyMan") %>' Width="35%" />
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
                <asp:Button ID="bbNew_Empty" runat="server" CausesValidation="false" CommandName="New" CssClass="button-Black" Text="新增" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CommandName="New" CssClass="button-Black" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CommandName="Edit" CssClass="button-Black" Text="修改" Width="120px" />
                &nbsp;<asp:Button ID="bbDelete_List" runat="server" CausesValidation="False" OnClick="bbDelete_List_Click" CssClass="button-Red" Text="刪除" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseDate_List" runat="server" CssClass="text-Right-Blue" Text="憑單及碼表日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColWidth-8Col">
                            <asp:Label ID="eCaseNo_List" runat="server" Text='<%# Eval("CaseNo") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCar_ID_List" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCar_ID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriver_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDriver_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver") %>' Width="35%" />
                            <asp:Label ID="eDriverName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="錯誤原因：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Height="95%" Width="98%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbInspector_List" runat="server" CssClass="text-Right-Blue" Text="查核人員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eInspector_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Inspector") %>' Width="35%" />
                            <asp:Label ID="eInspectorName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InspectorName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColBorder">
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
            <LocalReport ReportPath="Report\RunsheetDailyCheckP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsRunsheetDailyCheck_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, CaseDate, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = r.DepNo)) AS DepName, Car_ID, Driver, DriverName, Remark, Inspector, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = r.Inspector)) AS InspectorName, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = r.BuMan)) AS BuManName, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = r.ModifyMan)) AS ModifyManName FROM RunsheetDailyCheck AS r WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsRunsheetDailyCheck_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        DeleteCommand="DELETE FROM RunsheetDailyCheck WHERE (CaseNo = @CaseNo)"
        InsertCommand="INSERT INTO RunsheetDailyCheck(CaseNo, CaseDate, DepNo, Car_ID, Driver, DriverName, Remark, Inspector, BuDate, BuMan) VALUES (@CaseNo, @CaseDate, @DepNo, @Car_ID, @Driver, @DriverName, @Remark, @Inspector, @BuDate, @BuMan)"
        SelectCommand="SELECT CaseNo, CaseDate, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = r.DepNo)) AS DepName, Car_ID, Driver, DriverName, Remark, Inspector, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = r.Inspector)) AS InspectorName, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = r.BuMan)) AS BuManName, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = r.ModifyMan)) AS ModifyManName FROM RunsheetDailyCheck AS r WHERE (CaseNo = @CaseNo)"
        UpdateCommand="UPDATE RunsheetDailyCheck SET CaseDate = @CaseDate, DepNo = @DepNo, Car_ID = @Car_ID, Driver = @Driver, DriverName = @DriverName, Remark = @Remark, Inspector = @Inspector, ModifyDate = @ModifyDate, ModifyMan = @ModifyMan WHERE (CaseNo = @CaseNo)"
        OnDeleted="sdsRunsheetDailyCheck_Detail_Deleted"
        OnInserted="sdsRunsheetDailyCheck_Detail_Inserted"
        OnInserting="sdsRunsheetDailyCheck_Detail_Inserting"
        OnUpdated="sdsRunsheetDailyCheck_Detail_Updated"
        OnUpdating="sdsRunsheetDailyCheck_Detail_Updating">
        <DeleteParameters>
            <asp:Parameter Name="CaseNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="CaseDate" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="DriverName" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="Inspector" />
            <asp:Parameter Name="BuDate" />
            <asp:Parameter Name="BuMan" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridRunsheetDailyCheck_List" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="CaseDate" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="DriverName" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="Inspector" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="CaseNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
</asp:Content>
