<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="RescueCarAssignList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.RescueCarAssignList" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="RescueCarAssignListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">救援車勤務派遣登記表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbListYM_Search" runat="server" CssClass="text-Right-Blue" Text="派車年月：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eListYear_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:TextBox ID="eListMonth_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text=" 月" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="派車單位：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:DropDownList ID="ddlDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlDepNo_Search_SelectedIndexChanged" Width="95%"
                        DataSourceID="sdsDepNo_Search" DataTextField="ListTitle" DataValueField="DepNo" />
                    <asp:Label ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbAssignMan_Search" runat="server" CssClass="text-Right-Blue" Text="用車人 (含隨車人員)：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eAssignMan_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignMan_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eAssignManNAme_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" OnClick="bbSearch_Click" Text="查詢" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbPrintMonthList" runat="server" CssClass="button-Black" OnClick="bbPrintMonthList_Click" Text="預覽行車憑單" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbPrintPersonList" runat="server" CssClass="button-Black" OnClick="bbPrintPersonList_Click" Text="預覽人員公出記錄" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbPrintCarAssign" runat="server" CssClass="button-Black" OnClick="bbPrintCarAssign_Click" Text="預覽派車單" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" OnClick="bbClose_Click" Text="結束" Width="95%" />
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
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridRescueCarAssignList" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataKeyNames="ListNo" DataSourceID="sdsRescueCarAssignList" GridLines="None" Width="100%" PageSize="5">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="ListNo" HeaderText="ListNo" ReadOnly="True" SortExpression="ListNo" Visible="False" />
                <asp:BoundField DataField="CaseDate" DataFormatString="{0:D}" HeaderText="派車日期" SortExpression="CaseDate" />
                <asp:BoundField DataField="Car_ID" HeaderText="車號" SortExpression="Car_ID" />
                <asp:BoundField DataField="StartPosition" HeaderText="起點地名" SortExpression="StartPosition" />
                <asp:BoundField DataField="StartDate" DataFormatString="{0:D}" HeaderText="起點日期" SortExpression="StartDate" />
                <asp:BoundField DataField="StartTime" HeaderText="起點時刻" SortExpression="StartTime" />
                <asp:BoundField DataField="StartMeterMiles" HeaderText="起點路碼表" SortExpression="StartMeterMiles" />
                <asp:BoundField DataField="FinishPosition" HeaderText="終點地名" SortExpression="FinishPosition" />
                <asp:BoundField DataField="FinishDate" DataFormatString="{0:D}" HeaderText="終點日期" SortExpression="FinishDate" />
                <asp:BoundField DataField="FinishTime" HeaderText="終點時刻" SortExpression="FinishTime" />
                <asp:BoundField DataField="FinishMeterMiles" HeaderText="終點路碼表" SortExpression="FinishMeterMiles" />
                <asp:BoundField DataField="AssignManName" HeaderText="使用者" ReadOnly="True" SortExpression="AssignManName" />
                <asp:BoundField DataField="DepName" HeaderText="單位" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="FirstSupportManName" HeaderText="隨車人員1" ReadOnly="True" SortExpression="FirstSupportManName" />
                <asp:BoundField DataField="SecondSupportManName" HeaderText="隨車人員2" ReadOnly="True" SortExpression="SecondSupportManName" />
                <asp:BoundField DataField="AssignReason" HeaderText="申請事由" SortExpression="AssignReason" />
                <asp:BoundField DataField="ReasonNote" HeaderText="事項說明" SortExpression="ReasonNote" />
                <asp:BoundField DataField="UsedMiles" HeaderText="本次公里數" SortExpression="UsedMiles" />
                <asp:BoundField DataField="UsedMin" HeaderText="本日工時(分)" SortExpression="UsedMin" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" />
                <asp:BoundField DataField="BuManName" HeaderText="建檔人" ReadOnly="True" SortExpression="BuManName" Visible="False" />
                <asp:BoundField DataField="BuDate" DataFormatString="{0:D}" HeaderText="建檔日期" SortExpression="BuDate" Visible="False" />
                <asp:BoundField DataField="ModifyManName" HeaderText="異動人" ReadOnly="True" SortExpression="ModifyManName" Visible="False" />
                <asp:BoundField DataField="ModifyDate" DataFormatString="{0:D}" HeaderText="異動日期" SortExpression="ModifyDate" Visible="False" />
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
        <asp:FormView ID="fvRescueCarAssignDetail" runat="server" DataKeyNames="ListNo" DataSourceID="sdsRescueCarAssignDetail" Width="100%" OnDataBound="fvRescueCarAssignDetail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Black" OnClick="bbOK_Edit_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbCaseNo_Edit" runat="server" CssClass="text-Right-Blue" Text="憑單單號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eListNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ListNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbCaseDate_Edit" runat="server" CssClass="text-Right-Blue" Text="派車日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eCaseDate_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eCaseDate_Edit_TextChanged" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbCarID_Edit" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eCar_ID_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eCar_ID_Edit_TextChanged" Text='<%# Bind("Car_ID") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="6">
                                    <asp:Label ID="lbTitle1_Edit" runat="server" CssClass="titleText-Blue" Text="起點" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbStartPosition_Edit" runat="server" CssClass="text-Right-Blue" Text="地名：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eStartPosition_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("StartPosition") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbStartTime_Edit" runat="server" CssClass="text-Right-Blue" Text="日期時刻：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eStartDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StartDate","{0:yyyy/MM/dd}") %>' Width="55%" />
                                    <asp:TextBox ID="eStartTime_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eStartTime_Edit_TextChanged" Text='<%# Bind("StartTime") %>' Width="35%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbStartMeterMiles_Edit" runat="server" CssClass="text-Right-Blue" Text="路碼表：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eStartMeterMiles_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eStartMeterMiles_Edit_TextChanged" Text='<%# Bind("StartMeterMiles") %>' wid="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="6">
                                    <asp:Label ID="lbTitle2_Edit" runat="server" CssClass="titleText-Blue" Text="終點" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbFinishPosition_Edit" runat="server" CssClass="text-Right-Blue" Text="地點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eFinishPosition_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("FinishPosition") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbFinishTime_Edit" runat="server" CssClass="text-Right-Blue" Text="日期時刻：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eFinishDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FinishDate","{0:yyyy/MM/dd}") %>' Width="55%" />
                                    <asp:TextBox ID="eFinishTime_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eStartTime_Edit_TextChanged" Text='<%# Bind("FinishTime") %>' Width="35%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbFinishMeterMiles_Edit" runat="server" CssClass="text-Right-Blue" Text="路碼表：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eFinishMeterMiles_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eStartMeterMiles_Edit_TextChanged" Text='<%# Bind("FinishMeterMiles") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbAssignMan_Edit" runat="server" CssClass="text-Right-Blue" Text="使用人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eAssignMan_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignMan_Edit_TextChanged" Text='<%# Bind("AssignMan") %>' Width="35%" />
                                    <asp:Label ID="eAssignManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Text='<%# Bind("DepNo") %>' Width="35%" />
                                    <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbFirstSupportMan_Edit" runat="server" CssClass="text-Right-Blue" Text="隨車人員一：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eFirstSupportMan_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eFirstSupportMan_Edit_TextChanged" Text='<%# Bind("FirstSupportMan") %>' Width="35%" />
                                    <asp:Label ID="eFirstSupportManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FirstSupportManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbSecndSupportMan_Edit" runat="server" CssClass="text-Right-Blue" Text="隨車人員二：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eSecondSupportMan_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eSecondSupportMan_Edit_TextChanged" Text='<%# Bind("SecondSupportMan") %>' Width="35%" />
                                    <asp:Label ID="eSecondSupportManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SecondSupportManName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-6Col" rowspan="2">
                                    <asp:Label ID="lbAssignReason_Edit" runat="server" CssClass="text-Right-Blue" Text="申請事由：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="3">
                                    <asp:RadioButtonList ID="rbAssignReason_Edit" runat="server" CssClass="text-Left-Black" RepeatColumns="3" AutoPostBack="true"
                                        OnSelectedIndexChanged="rbAssignReason_Edit_SelectedIndexChanged" Width="95%">
                                        <asp:ListItem Value="救援" Text="救援" Selected="True" />
                                        <asp:ListItem Value="領料" Text="領料" />
                                        <asp:ListItem Value="其他" Text="其他" />
                                    </asp:RadioButtonList>
                                    <asp:Label ID="eAssignReason_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("AssignReason") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbTargetCarID_Edit" runat="server" CssClass="text-Right-Blue" Text="救援車號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eTargetCarID_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("TargetCarID") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                                    <asp:TextBox ID="eReasonNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("ReasonNote") %>' Width="98%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Width="98%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColWidth-6Col">
                                    <asp:Label ID="eUsedMiles_Edit" runat="server" Text='<%# Bind("UsedMiles") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColWidth-6Col">
                                    <asp:Label ID="eUsedMin_Edit" runat="server" Text='<%# Bind("UsedMin") %>' Visible="false" />
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
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Insert" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbCaseNo_INS" runat="server" CssClass="text-Right-Blue" Text="憑單單號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eListNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ListNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbCaseDate_INS" runat="server" CssClass="text-Right-Blue" Text="派車日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eCaseDate_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eCaseDate_INS_TextChanged" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbCarID_INS" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eCar_ID_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eCar_ID_INS_TextChanged" Text='<%# Bind("Car_ID") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="6">
                                    <asp:Label ID="lbTitle1_INS" runat="server" CssClass="titleText-Blue" Text="起點" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbStartPosition_INS" runat="server" CssClass="text-Right-Blue" Text="地名：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eStartPosition_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("StartPosition") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbStartTime_INS" runat="server" CssClass="text-Right-Blue" Text="日期時刻：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eStartDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StartDate","{0:yyyy/MM/dd}") %>' Width="55%" />
                                    <asp:TextBox ID="eStartTime_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eStartTime_INS_TextChanged" Text='<%# Bind("StartTime") %>' Width="35%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbStartMeterMiles_INS" runat="server" CssClass="text-Right-Blue" Text="路碼表：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eStartMeterMiles_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eStartMeterMiles_INS_TextChanged" Text='<%# Bind("StartMeterMiles") %>' wid="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="6">
                                    <asp:Label ID="lbTitle2_INS" runat="server" CssClass="titleText-Blue" Text="終點" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbFinishPosition_INS" runat="server" CssClass="text-Right-Blue" Text="地點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eFinishPosition_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("FinishPosition") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbFinishTime_INS" runat="server" CssClass="text-Right-Blue" Text="日期時刻：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eFinishDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FinishDate","{0:yyyy/MM/dd}") %>' Width="55%" />
                                    <asp:TextBox ID="eFinishTime_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eStartTime_INS_TextChanged" Text='<%# Bind("FinishTime") %>' Width="35%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbFinishMeterMiles_INS" runat="server" CssClass="text-Right-Blue" Text="路碼表：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eFinishMeterMiles_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eStartMeterMiles_INS_TextChanged" Text='<%# Bind("FinishMeterMiles") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbAssignMan_INS" runat="server" CssClass="text-Right-Blue" Text="使用人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eAssignMan_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignMan_INS_TextChanged" Text='<%# Bind("AssignMan") %>' Width="35%" />
                                    <asp:Label ID="eAssignManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Bind("DepNo") %>' Width="35%" />
                                    <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbFirstSupportMan_INS" runat="server" CssClass="text-Right-Blue" Text="隨車人員一：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eFirstSupportMan_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eFirstSupportMan_INS_TextChanged" Text='<%# Bind("FirstSupportMan") %>' Width="35%" />
                                    <asp:Label ID="eFirstSupportManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FirstSupportManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbSecndSupportMan_INS" runat="server" CssClass="text-Right-Blue" Text="隨車人員二：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eSecondSupportMan_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eSecondSupportMan_INS_TextChanged" Text='<%# Bind("SecondSupportMan") %>' Width="35%" />
                                    <asp:Label ID="eSecondSupportManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SecondSupportManName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-6Col" rowspan="2">
                                    <asp:Label ID="lbAssignReason_INS" runat="server" CssClass="text-Right-Blue" Text="申請事由：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="3">
                                    <asp:RadioButtonList ID="rbAssignReason_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnSelectedIndexChanged="rbAssignReason_INS_SelectedIndexChanged" RepeatColumns="3" Width="95%">
                                        <asp:ListItem Value="救援" Text="救援" Selected="True" />
                                        <asp:ListItem Value="領料" Text="領料" />
                                        <asp:ListItem Value="其他" Text="其他" />
                                    </asp:RadioButtonList>
                                    <asp:TextBox ID="eAssignReason_INS" runat="server" Text='<%# Bind("AssignReason") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbTargetCarID_INS" runat="server" CssClass="text-Right-Blue" Text="救援車號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eTargetCarID_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("TargetCarID") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                                    <asp:TextBox ID="eReasonNote_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("ReasonNote") %>' Width="98%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Width="98%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColWidth-6Col">
                                    <asp:Label ID="eUsedMiles_INS" runat="server" Text='<%# Bind("UsedMiles") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                                    <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColWidth-6Col">
                                    <asp:Label ID="eUsedMin_INS" runat="server" Text='<%# Bind("UsedMin") %>' Visible="false" />
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
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="Edit" Text="編輯" Width="120px" />
                &nbsp;<asp:Button ID="bbDelete_List" runat="server" CausesValidation="False" CssClass="button-Red" OnClick="bbDelete_List_Click" Text="刪除" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseNo_List" runat="server" CssClass="text-Right-Blue" Text="憑單單號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eListNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ListNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseDate_List" runat="server" CssClass="text-Right-Blue" Text="派車日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCarID_List" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCar_ID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="6">
                            <asp:Label ID="lbTitle1_List" runat="server" CssClass="titleText-Blue" Text="起點" Width="98%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbStartPosition_List" runat="server" CssClass="text-Right-Blue" Text="地名：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eStartPosition_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StartPosition") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbStartTime_List" runat="server" CssClass="text-Right-Blue" Text="日期時刻：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eStartDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StartDate","{0:yyyy/MM/dd}") %>' Width="55%" />
                            <asp:Label ID="eStartTime_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StartTime") %>' Width="35%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbStartMeterMiles_List" runat="server" CssClass="text-Right-Blue" Text="路碼表：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eStartMeterMiles_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StartMeterMiles") %>' wid="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="6">
                            <asp:Label ID="lbTitle2_List" runat="server" CssClass="titleText-Blue" Text="終點" Width="98%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbFinishPosition_List" runat="server" CssClass="text-Right-Blue" Text="地點：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eFinishPosition_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FinishPosition") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbFinishTime_List" runat="server" CssClass="text-Right-Blue" Text="日期時刻：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eFinishDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FinishDate","{0:yyyy/MM/dd}") %>' Width="55%" />
                            <asp:Label ID="eFinishTime_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FinishTime") %>' Width="35%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbFinishMeterMiles_List" runat="server" CssClass="text-Right-Blue" Text="路碼表：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eFinishMeterMiles_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FinishMeterMiles") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAssignMan_List" runat="server" CssClass="text-Right-Blue" Text="使用人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                            <asp:Label ID="eAssignMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' Width="35%" />
                            <asp:Label ID="eAssignManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbFirstSupportMan_List" runat="server" CssClass="text-Right-Blue" Text="隨車人員一：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                            <asp:Label ID="eFirstSupportMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FirstSupportMan") %>' Width="35%" />
                            <asp:Label ID="eFirstSupportManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FirstSupportManName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbSecndSupportMan_List" runat="server" CssClass="text-Right-Blue" Text="隨車人員二：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                            <asp:Label ID="eSecondSupportMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SecondSupportMan") %>' Width="35%" />
                            <asp:Label ID="eSecondSupportManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SecondSupportManName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" rowspan="2">
                            <asp:Label ID="lbAssignReason_List" runat="server" CssClass="text-Right-Blue" Text="申請事由：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="3">
                            <asp:Label ID="eAssignReason_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignReason") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbTargetCarID_List" runat="server" CssClass="text-Right-Blue" Text="救援車號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eTargetCarID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TargetCarID") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eReasonNote_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("ReasonNote") %>' Width="98%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Width="98%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                            <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColWidth-6Col">
                            <asp:Label ID="eUsedMiles_List" runat="server" Text='<%# Eval("UsedMiles") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                            <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColWidth-6Col">
                            <asp:Label ID="eUsedMin_List" runat="server" Text='<%# Eval("UsedMin") %>' Visible="false" />
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
            <LocalReport ReportPath="Report\RescueCarAssignListP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsDepNo_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST('' AS nvarchar) AS ListTitle, CAST('' AS nvarchar) AS DepNo UNION ALL SELECT DISTINCT DepNo + (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = r.DepNo)) AS ListTitle, DepNo FROM RescueCarAssignList AS r"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsRescueCarAssignList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT ListNo, CaseDate, Car_ID, StartPosition, StartTime, StartMeterMiles, FinishPosition, FinishTime, FinishMeterMiles, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = r.AssignMan)) AS AssignManName, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = r.DepNo)) AS DepName, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_4 WHERE (EMPNO = r.FirstSupportMan)) AS FirstSupportManName, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (EMPNO = r.SecondSupportMan)) AS SecondSupportManName, AssignReason, ReasonNote, UsedMiles, UsedMin, Remark, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = r.BuMan)) AS BuManName, BuDate, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = r.ModifyMan)) AS ModifyManName, ModifyDate, TargetCarID, StartDate, FinishDate FROM RescueCarAssignList AS r WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsRescueCarAssignDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        DeleteCommand="DELETE FROM RescueCarAssignList WHERE (ListNo = @ListNo)"
        InsertCommand="INSERT INTO RescueCarAssignList(ListNo, CaseDate, Car_ID, StartPosition, StartDate, StartTime, StartMeterMiles, FinishPosition, FinishDate, FinishTime, FinishMeterMiles, AssignMan, DepNo, FirstSupportMan, SecondSupportMan, AssignReason, ReasonNote, UsedMiles, UsedMin, Remark, BuMan, BuDate, TargetCarID) VALUES (@ListNo, @CaseDate, @Car_ID, @StartPosition, @StartDate, @StartTime, @StartMeterMiles, @FinishPosition, @FinishDate, @FinishTime, @FinishMeterMiles, @AssignMan, @DepNo, @FirstSupportMan, @SecondSupportMan, @AssignReason, @ReasonNote, @UsedMiles, @UsedMin, @Remark, @BuMan, @BuDate, @TargetCarID)"
        SelectCommand="SELECT ListNo, CaseDate, Car_ID, StartPosition, StartDate, StartTime, StartMeterMiles, FinishPosition, FinishDate, FinishTime, FinishMeterMiles, AssignMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = r.AssignMan)) AS AssignManName, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = r.DepNo)) AS DepName, FirstSupportMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_4 WHERE (EMPNO = r.FirstSupportMan)) AS FirstSupportManName, SecondSupportMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (EMPNO = r.SecondSupportMan)) AS SecondSupportManName, AssignReason, ReasonNote, UsedMiles, UsedMin, Remark, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = r.BuMan)) AS BuManName, BuDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = r.ModifyMan)) AS ModifyManName, ModifyDate, TargetCarID FROM RescueCarAssignList AS r WHERE (ListNo = @ListNo)"
        UpdateCommand="UPDATE RescueCarAssignList SET CaseDate = @CaseDate, Car_ID = @Car_ID, StartPosition = @StartPosition, StartDate = @StartDate, StartTime = @StartTime, StartMeterMiles = @StartMeterMiles, FinishPosition = @FinishPosition, FinishDate = @FinishDate, FinishTime = @FinishTime, FinishMeterMiles = @FinishMeterMiles, AssignMan = @AssignMan, DepNo = @DepNo, FirstSupportMan = @FirstSupportMan, SecondSupportMan = @SecondSupportMan, AssignReason = @AssignReason, ReasonNote = @ReasonNote, UsedMiles = @UsedMiles, UsedMin = @UsedMin, Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate, TargetCarID = @TargetCarID WHERE (ListNo = @ListNo)"
        OnDeleted="sdsRescueCarAssignDetail_Deleted"
        OnInserted="sdsRescueCarAssignDetail_Inserted"
        OnInserting="sdsRescueCarAssignDetail_Inserting">
        <DeleteParameters>
            <asp:Parameter Name="ListNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="ListNo" />
            <asp:Parameter Name="CaseDate" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="StartPosition" />
            <asp:Parameter Name="StartTime" />
            <asp:Parameter Name="StartMeterMiles" />
            <asp:Parameter Name="FinishPosition" />
            <asp:Parameter Name="FinishTime" />
            <asp:Parameter Name="FinishMeterMiles" />
            <asp:Parameter Name="AssignMan" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="FirstSupportMan" />
            <asp:Parameter Name="SecondSupportMan" />
            <asp:Parameter Name="AssignReason" />
            <asp:Parameter Name="ReasonNote" />
            <asp:Parameter Name="UsedMiles" />
            <asp:Parameter Name="UsedMin" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="BuDate" />
            <asp:Parameter Name="TargetCarID" />
            <asp:Parameter Name="StartDate" />
            <asp:Parameter Name="FinishDate" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridRescueCarAssignList" Name="ListNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="CaseDate" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="StartPosition" />
            <asp:Parameter Name="StartTime" />
            <asp:Parameter Name="StartMeterMiles" />
            <asp:Parameter Name="FinishPosition" />
            <asp:Parameter Name="FinishTime" />
            <asp:Parameter Name="FinishMeterMiles" />
            <asp:Parameter Name="AssignMan" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="FirstSupportMan" />
            <asp:Parameter Name="SecondSupportMan" />
            <asp:Parameter Name="AssignReason" />
            <asp:Parameter Name="ReasonNote" />
            <asp:Parameter Name="UsedMiles" />
            <asp:Parameter Name="UsedMin" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="ListNo" />
            <asp:Parameter Name="TargetCarID" />
            <asp:Parameter Name="StartDate" />
            <asp:Parameter Name="FinishDate" />
        </UpdateParameters>
    </asp:SqlDataSource>
</asp:Content>
