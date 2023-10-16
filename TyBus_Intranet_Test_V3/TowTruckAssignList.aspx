<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="TowTruckAssignList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.TowTruckAssignList" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="TowTruckAssignListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">拖吊車叫用記錄表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCaseDate_Search" runat="server" CssClass="text-Right-Blue" Text="拖吊日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eCaseDate_S_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eCaseDate_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Car" runat="server" CssClass="text-Right-Blue" Text="車輛所屬站別：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Car_S_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Car_S_Search_TextChanged" Width="40%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eDepNo_Car_E_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Car_E_Search_TextChanged" Width="40%" />
                    <br />
                    <asp:Label ID="eDepName_Car_S_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="eDepName_Car_E_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Man_Search" runat="server" CssClass="text-Right-Blue" Text="檢修人員單位：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eDepNo_Man_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Man_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eDepName_Man_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbAssignMan_Search" runat="server" CssClass="text-Right-Blue" Text="檢修人員：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eAssignMan_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignMan_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eAssignName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCarID_Search" runat="server" CssClass="text-Right-Blue" Text="車號" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCarID_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDriver_Search" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eDriver_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eDriverName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col" colspan="8">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" OnClick="bbSearch_Click" Text="查詢" Width="120px" />
                    <asp:Button ID="bbPrintList" runat="server" CssClass="button-Red" OnClick="bbPrintList_Click" Text="拖吊車紀錄清冊" />
                    <asp:Button ID="bbPrintDetail" runat="server" CssClass="button-Black" OnClick="bbPrintDetail_Click" Text="拖吊內容紀錄表" />
                    <asp:Button ID="bbPrintCostList" runat="server" CssClass="button-Blue" OnClick="bbPrintCostList_Click" Text="拖吊維修費用月統計表" />
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
        <asp:GridView ID="gridTowTruckAssignList" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1"
            DataKeyNames="CaseNo" DataSourceID="sdsTowTruckAssignList" GridLines="None" Width="100%">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="CaseNo" HeaderText="CaseNo" ReadOnly="True" SortExpression="CaseNo" Visible="false" />
                <asp:BoundField DataField="CaseType_C" HeaderText="類別" SortExpression="CaseType_C" />
                <asp:BoundField DataField="CaseDate" DataFormatString="{0:D}" HeaderText="發生日期" SortExpression="CaseDate" />
                <asp:BoundField DataField="CaseTime" HeaderText="發生時間" SortExpression="CaseTime" />
                <asp:BoundField DataField="Car_ID" HeaderText="車號" SortExpression="Car_ID" />
                <asp:BoundField DataField="DepName_Car" HeaderText="車輛站別" ReadOnly="True" SortExpression="DepName_Car" />
                <asp:BoundField DataField="CasePosition" HeaderText="發生地點" SortExpression="CasePosition" />
                <asp:BoundField DataField="ParkingPosition" HeaderText="停放地點" SortExpression="ParkingPosition" />
                <asp:BoundField DataField="AssignMan" HeaderText="AssignMan" SortExpression="AssignMan" Visible="False" />
                <asp:BoundField DataField="AssignManName" HeaderText="處理人員" ReadOnly="True" SortExpression="AssignManName" />
                <asp:BoundField DataField="FirstSupportMan" HeaderText="FirstSupportMan" SortExpression="FirstSupportMan" Visible="False" />
                <asp:BoundField DataField="FirstSupportManName" HeaderText="隨行處理人員1" ReadOnly="True" SortExpression="FirstSupportManName" />
                <asp:BoundField DataField="SecondSupportMan" HeaderText="SecondSupportMan" SortExpression="SecondSupportMan" Visible="False" />
                <asp:BoundField DataField="SecondSupportManNAme" HeaderText="隨行處理人員2" ReadOnly="True" SortExpression="SecondSupportManNAme" />
                <asp:BoundField DataField="Determination" HeaderText="判別" SortExpression="Determination" />
                <asp:BoundField DataField="IsChecked" HeaderText="主管已確認" SortExpression="IsChecked" />
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
        <asp:FormView ID="fvTowTruckAssignDetail" runat="server" DataKeyNames="CaseNo" DataSourceID="sdsTowTruckAssignDetail" Width="100%" OnDataBound="fvTowTruckAssignDetail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Black" OnClick="bbOK_Edit_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseNo_Edit" runat="server" CssClass="text-Right-Blue" Text="拖吊單號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eCaseNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseType_Edit" runat="server" CssClass="text-Right-Blue" Text="類別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:DropDownList ID="ddlCaseType_Edit" runat="server" CssClass="text-Left-Black"
                                        AutoPostBack="true" OnSelectedIndexChanged="ddlCaseType_Edit_SelectedIndexChanged" Width="95%">
                                    </asp:DropDownList>
                                    <asp:Label ID="eCaseType_Edit" runat="server" Text='<%# Eval("CaseType") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="cbIsChecked_Edit" runat="server" CssClass="text-Left-Black" Text="主管已確認" Enabled="false" OnCheckedChanged="cbIsChecked_Edit_CheckedChanged" />
                                    <asp:Label ID="eIsChecked_Edit" runat="server" Text='<%# Eval("IsChecked") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarID_Edit" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eCar_ID_Edit_TextChanged" Text='<%# Eval("Car_ID") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Car_Edit" runat="server" CssClass="text-Right-Blue" Text="車輛所屬站別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_Car_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Car_Edit_TextChanged" Text='<%# Eval("DepNo_Car") %>' Width="35%" />
                                    <asp:Label ID="eDepName_Car_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName_Car") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDriver_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Edit_TextChanged" Text='<%# Eval("Driver") %>' Width="35%" />
                                    <asp:Label ID="eDriverName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCar_Class_Edit" runat="server" CssClass="text-Right-Blue" Text="廠牌：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCar_ClassName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ClassName") %>' Width="95%" />
                                    <asp:Label ID="eCar_Class_Edit" runat="server" Text='<%# Eval("Car_Class") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbGetLicDate_Edit" runat="server" CssClass="text-Right-Blue" Text="領牌日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eGetLicDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("getlicdate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarAge_Edit" runat="server" CssClass="text-Right-Blue" Text="車齡：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarAge_Year_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarAge_Year") %>' />
                                    <asp:Label ID="lbSplit1_Edit" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                                    <asp:Label ID="eCarAge_Month_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarAge_Month") %>' />
                                    <asp:Label ID="lbSplit2_Edit" runat="server" CssClass="text-Left-Black" Text=" 月" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPoint_Edit" runat="server" CssClass="text-Right-Blue" Text="車輛用途：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="ePointName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PointName") %>' Width="95%" />
                                    <asp:Label ID="ePoint_Edit" runat="server" Text='<%# Eval("point") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLastServiceDate_Edit" runat="server" CssClass="text-Right-Blue" Text="最後保養日：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLastServiceDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastServiceDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarKMs_Edit" runat="server" CssClass="text-Right-Blue" Text="里程數：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCarKMs_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarKMs") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseDate_Edit" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCaseDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseTime_Edit" runat="server" CssClass="text-Right-Blue" Text="發生時間：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCaseTime_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseTime") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCasePosition_Edit" runat="server" CssClass="text-Right-Blue" Text="發生地點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eCasePosition_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CasePosition") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbParkingPosition_Edit" runat="server" CssClass="text-Right-Blue" Text="停放地點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eParkingPosition_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ParkingPosition") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTowingCost_Edit" runat="server" CssClass="text-Right-Blue" Text="拖吊費用：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTowingCost_Edit" runat="server" CssClass="text-Left-Black" Text='<%#Eval("TowingCost") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDetermination_Edit" runat="server" CssClass="text-Right-Blue" Text="判別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eDetermination_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Determination") %>' Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDispose_Edit" runat="server" CssClass="text-Right-Blue" Text="處理方式：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDispose_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Dispose") %>' Height="95%" Width="98%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignMan_Edit" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignMan_Edit" runat="server" CssClass="text-Left-Black"
                                        AutoPostBack="true" OnTextChanged="eAssignMan_Edit_TextChanged" Text='<%# Eval("AssignMan") %>' Width="95%" />
                                    <br />
                                    <asp:Label ID="eAssignManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFirstSupportMan_Edit" runat="server" CssClass="text-Right-Blue" Text="隨行處理人員1：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eFirstSupportMan_Edit" runat="server" CssClass="text-Left-Black"
                                        AutoPostBack="true" OnTextChanged="eFirstSupportMan_Edit_TextChanged" Text='<%# Eval("FirstSupportMan") %>' Width="95%" />
                                    <br />
                                    <asp:Label ID="eFirstSupportManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FirstSupportManName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSecondSupportMan_Edit" runat="server" CssClass="text-Right-Blue" Text="隨行處理人員2：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eSecondSupportMan_Edit" runat="server" CssClass="text-Left-Black"
                                        AutoPostBack="true" OnTextChanged="eSecondSupportMan_Edit_TextChanged" Text='<%# Eval("SecondSupportMan") %>' Width="95%" />
                                    <br />
                                    <asp:Label ID="eSecondSupportManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SecondSupportManName") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFaultParts_Edit" runat="server" CssClass="text-Right-Blue" Text="損壞部位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eFaultParts_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FaultParts") %>' Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFaultReason_Edit" runat="server" CssClass="text-Right-Blue" Text="損壞原因：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eFaultReason_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("FaultReason") %>' Height="95%" Width="98%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDamageAnalysis_Edit" runat="server" CssClass="text-Right-Blue" Text="損壞分析：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlDamageAnalysis_Edit" runat="server" CssClass="text-Left-Black"
                                        AutoPostBack="true" OnSelectedIndexChanged="ddlDamageAnalysis_Edit_SelectedIndexChanged" Width="95%">
                                    </asp:DropDownList>
                                    <asp:Label ID="eDamageAnalysis_Edit" runat="server" Text='<%#Eval("DamageAnalysis") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbWorkSheetNo_Edit" runat="server" CssClass="text-Right-Blue" Text="工作單號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eWorkSheetNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%#Eval("WorkSheetNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFixFee_Edit" runat="server" CssClass="text-Right-Blue" Text="維修費用：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eFixFee_Edit" runat="server" CssClass="text-Left-Black" Text='<%#Eval("FixFee") %>' Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFollowUp_Edit" runat="server" CssClass="text-Right-Blue" Text="後續追蹤：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eFollowUp_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("FollowUp") %>' Height="95%" Width="98%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbImprovements_Edit" runat="server" CssClass="text-Right-Blue" Text="改善措施：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eImprovements_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Improvements") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eDepNo_Man_Edit" runat="server" Text='<%# Eval("DepNo_Man") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
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
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Black" OnClick="bbOK_INS_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseNo_INS" runat="server" CssClass="text-Right-Blue" Text="拖吊單號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eCaseNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseType_INS" runat="server" CssClass="text-Right-Blue" Text="類別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:DropDownList ID="ddlCaseType_INS" runat="server" CssClass="text-Left-Black"
                                        AutoPostBack="true" OnSelectedIndexChanged="ddlCaseType_INS_SelectedIndexChanged" Width="95%">
                                    </asp:DropDownList>
                                    <asp:Label ID="eCaseType_INS" runat="server" Text='<%# Eval("CaseType") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="cbIsChecked_INS" runat="server" CssClass="text-Left-Black" Text="主管已確認" Enabled="false" OnCheckedChanged="cbIsChecked_INS_CheckedChanged" />
                                    <asp:Label ID="eIsChecked_INS" runat="server" Text='<%# Eval("IsChecked") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarID_INS" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eCar_ID_INS_TextChanged" Text='<%# Eval("Car_ID") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Car_INS" runat="server" CssClass="text-Right-Blue" Text="車輛所屬站別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_Car_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Car_INS_TextChanged" Text='<%# Eval("DepNo_Car") %>' Width="35%" />
                                    <asp:Label ID="eDepName_Car_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName_Car") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_INS" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDriver_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_INS_TextChanged" Text='<%# Eval("Driver") %>' Width="35%" />
                                    <asp:Label ID="eDriverName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCar_Class_INS" runat="server" CssClass="text-Right-Blue" Text="廠牌：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCar_ClassName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ClassName") %>' Width="95%" />
                                    <asp:Label ID="eCar_Class_INS" runat="server" Text='<%# Eval("Car_Class") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbGetLicDate_INS" runat="server" CssClass="text-Right-Blue" Text="領牌日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eGetLicDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("getlicdate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarAge_INS" runat="server" CssClass="text-Right-Blue" Text="車齡：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarAge_Year_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarAge_Year") %>' />
                                    <asp:Label ID="lbSplit1_INS" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                                    <asp:Label ID="eCarAge_Month_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarAge_Month") %>' />
                                    <asp:Label ID="lbSplit2_INS" runat="server" CssClass="text-Left-Black" Text=" 月" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPoint_INS" runat="server" CssClass="text-Right-Blue" Text="車輛用途：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="ePointName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PointName") %>' Width="95%" />
                                    <asp:Label ID="ePoint_INS" runat="server" Text='<%# Eval("point") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLastServiceDate_INS" runat="server" CssClass="text-Right-Blue" Text="最後保養日：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLastServiceDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastServiceDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarKMs_INS" runat="server" CssClass="text-Right-Blue" Text="里程數：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCarKMs_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarKMs") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseDate_INS" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCaseDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseTime_INS" runat="server" CssClass="text-Right-Blue" Text="發生時間：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCaseTime_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseTime") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCasePosition_INS" runat="server" CssClass="text-Right-Blue" Text="發生地點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eCasePosition_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CasePosition") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbParkingPosition_INS" runat="server" CssClass="text-Right-Blue" Text="停放地點：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eParkingPosition_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ParkingPosition") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTowingCost_INS" runat="server" CssClass="text-Right-Blue" Text="拖吊費用：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTowingCost_INS" runat="server" CssClass="text-Left-Black" Text='<%#Eval("TowingCost") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDetermination_INS" runat="server" CssClass="text-Right-Blue" Text="判別：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eDetermination_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Determination") %>' Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDispose_INS" runat="server" CssClass="text-Right-Blue" Text="處理方式：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDispose_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Dispose") %>' Height="95%" Width="98%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignMan_INS" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignMan_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignMan_INS_TextChanged" Text='<%# Eval("AssignMan") %>' Width="35%" />
                                    <br />
                                    <asp:Label ID="eAssignManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFirstSupportMan_INS" runat="server" CssClass="text-Right-Blue" Text="隨行處理人員1：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eFirstSupportMan_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eFirstSupportMan_INS_TextChanged" Text='<%# Eval("FirstSupportMan") %>' Width="35%" />
                                    <br />
                                    <asp:Label ID="eFirstSupportManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FirstSupportManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSecondSupportMan_INS" runat="server" CssClass="text-Right-Blue" Text="隨行處理人員2：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eSecondSupportMan_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eSecondSupportMan_INS_TextChanged" Text='<%# Eval("SecondSupportMan") %>' Width="35%" />
                                    <br />
                                    <asp:Label ID="eSecondSupportManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SecondSupportManNAme") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFaultParts_INS" runat="server" CssClass="text-Right-Blue" Text="損壞部位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eFaultParts_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FaultParts") %>' Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFaultReason_INS" runat="server" CssClass="text-Right-Blue" Text="損壞原因：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eFaultReason_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("FaultReason") %>' Height="95%" Width="98%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDamageAnalysis_INS" runat="server" CssClass="text-Right-Blue" Text="損壞分析：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlDamageAnalysis_INS" runat="server" CssClass="text-Left-Black"
                                        AutoPostBack="true" OnSelectedIndexChanged="ddlDamageAnalysis_INS_SelectedIndexChanged" Width="95%">
                                    </asp:DropDownList>
                                    <asp:Label ID="eDamageAnalysis_INS" runat="server" Text='<%#Eval("DamageAnalysis") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbWorkSheetNo_INS" runat="server" CssClass="text-Right-Blue" Text="工作單號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eWorkSheetNo_INS" runat="server" CssClass="text-Left-Black" Text='<%#Eval("WorkSheetNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFixFee_INS" runat="server" CssClass="text-Right-Blue" Text="維修費用：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eFixFee_INS" runat="server" CssClass="text-Left-Black" Text='<%#Eval("FixFee") %>' Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFollowUp_INS" runat="server" CssClass="text-Right-Blue" Text="後續追蹤：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eFollowUp_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("FollowUp") %>' Height="95%" Width="98%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbImprovements_INS" runat="server" CssClass="text-Right-Blue" Text="改善措施：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eImprovements_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Improvements") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eDepNo_Man_INS" runat="server" Text='<%# Eval("DepNo_Man") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
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
                <asp:Button ID="bbNew_Empty" runat="server" CausesValidation="false" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="Edit" Text="修改" Width="120px" />
                &nbsp;<asp:Button ID="bbDelete_List" runat="server" CausesValidation="False" CssClass="button-Red" OnClick="bbDelete_List_Click" Text="刪除" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseNo_List" runat="server" CssClass="text-Right-Blue" Text="拖吊單號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eCaseNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseType_List" runat="server" CssClass="text-Right-Blue" Text="類別：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eCaseType_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseType_C") %>' Width="95%" />
                            <asp:Label ID="eCaseType_List" runat="server" Text='<%# Eval("CaseType") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:CheckBox ID="cbIsChecked_List" runat="server" CssClass="text-Left-Black" Text="主管已確認" Enabled="false" />
                            <asp:Label ID="eIsChecked_List" runat="server" Text='<%# Eval("IsChecked") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCarID_List" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCar_ID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_Car_List" runat="server" CssClass="text-Right-Blue" Text="車輛所屬站別：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_Car_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo_Car") %>' Width="35%" />
                            <asp:Label ID="eDepName_Car_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName_Car") %>' Width="55%" />
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
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCar_Class_List" runat="server" CssClass="text-Right-Blue" Text="廠牌：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCar_Class_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_Class") %>' Visible="false" />
                            <asp:Label ID="eCar_ClassName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ClassName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbGetLicDate_List" runat="server" CssClass="text-Right-Blue" Text="領牌日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eGetLicDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("getlicdate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCarAge_List" runat="server" CssClass="text-Right-Blue" Text="車齡：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCarAge_Year_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarAge_Year") %>' />
                            <asp:Label ID="lbSplit1_List" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                            <asp:Label ID="eCarAge_Month_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarAge_Month") %>' />
                            <asp:Label ID="lbSplit2_List" runat="server" CssClass="text-Left-Black" Text=" 月" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPoint_List" runat="server" CssClass="text-Right-Blue" Text="車輛用途：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="ePoint_List" runat="server" Text='<%# Eval("point") %>' Visible="false" />
                            <asp:Label ID="ePointName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PointName") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLastServiceDate_List" runat="server" CssClass="text-Right-Blue" Text="最後保養日：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLastServiceDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastServiceDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCarKMs_List" runat="server" CssClass="text-Right-Blue" Text="里程數：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCarKMs_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarKMs") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseDate_List" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseTime_List" runat="server" CssClass="text-Right-Blue" Text="發生時間：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseTime_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseTime") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCasePosition_List" runat="server" CssClass="text-Right-Blue" Text="發生地點：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eCasePosition_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CasePosition") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbParkingPosition_List" runat="server" CssClass="text-Right-Blue" Text="停放地點：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eParkingPosition_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ParkingPosition") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbTowingCost_List" runat="server" CssClass="text-Right-Blue" Text="拖吊費用：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eTowingCost_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("TowingCost") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDetermination_List" runat="server" CssClass="text-Right-Blue" Text="判別：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="7">
                            <asp:Label ID="eDetermination_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Determination") %>' Width="98%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDispose_List" runat="server" CssClass="text-Right-Blue" Text="處理方式：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eDispose_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Dispose") %>' Height="95%" Width="98%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignMan_List" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAssignMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' Width="95%" />
                            <br />
                            <asp:Label ID="eAssignManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbFirstSupportMan_List" runat="server" CssClass="text-Right-Blue" Text="隨行處理人員1：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eFirstSupportMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FirstSupportMan") %>' Width="95%" />
                            <br />
                            <asp:Label ID="eFirstSupportManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FirstSupportManName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSecondSupportMan_List" runat="server" CssClass="text-Right-Blue" Text="隨行處理人員2：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSecondSupportMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SecondSupportMan") %>' Width="95%" />
                            <br />
                            <asp:Label ID="eSecondSupportManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("SecondSupportManNAme") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbFaultParts_List" runat="server" CssClass="text-Right-Blue" Text="損壞部位：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eFaultParts_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FaultParts") %>' Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbFaultReason_List" runat="server" CssClass="text-Right-Blue" Text="損壞原因：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="3">
                            <asp:TextBox ID="eFaultReason_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("FaultReason") %>' Height="95%" Width="98%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDamageAnalysis_List" runat="server" CssClass="text-Right-Blue" Text="損壞分析：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDamageAnalysis_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("DamageAnalysis_C") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbWorkSheetNo_List" runat="server" CssClass="text-Right-Blue" Text="工作單號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eWorkSheetNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("WorkSheetNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbFixFee_List" runat="server" CssClass="text-Right-Blue" Text="維修費用：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eFixFee_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FixFee") %>' Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbFollowUp_List" runat="server" CssClass="text-Right-Blue" Text="後續追蹤：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eFollowUp_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("FollowUp") %>' Height="95%" Width="98%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbImprovements_List" runat="server" CssClass="text-Right-Blue" Text="改善措施：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eImprovements_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Improvements") %>' Height="95%" Width="98%" />
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
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_Man_List" runat="server" Text='<%# Eval("DepNo_Man") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2" />
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
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
            <LocalReport ReportPath="Report\TowTruckAssignRecordP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsTowTruckAssignList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '拖吊車叫用記錄表TowTruckAssign  CaseType') AND (CLASSNO = t.CaseType)) AS CaseType_C, CaseNo, CaseDate, Car_ID, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = t.DepNo_Car)) AS DepName_Car, CasePosition, ParkingPosition, CaseTime, AssignMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = t.AssignMan)) AS AssignManName, FirstSupportMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = t.FirstSupportMan)) AS FirstSupportManName, SecondSupportMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = t.SecondSupportMan)) AS SecondSupportManNAme, Determination, IsChecked FROM TowTruckAssignList AS t WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsTowTruckAssignDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT t.CaseNo, t.CaseDate, t.Car_ID, t.Driver, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = t.Driver)) AS DriverName, t.DepNo_Car, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = t.DepNo_Car)) AS DepName_Car, c.Car_Class, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = c.Car_Class) AND (FKEY = '車輛資料作業    Car_infoA       CAR_CLASS')) AS Car_ClassName, c.point, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (CLASSNO = c.point) AND (FKEY = '車輛資料作業    Car_infoA       POINT')) AS PointName, DATEDIFF(Month, c.getlicdate, GETDATE()) / 12 AS CarAge_Year, DATEDIFF(Month, c.getlicdate, GETDATE()) % 12 AS CarAge_Month, c.getlicdate, t.CasePosition, t.ParkingPosition, t.CaseTime, t.AssignMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_5 WHERE (EMPNO = t.AssignMan)) AS AssignManName, t.DepNo_Man, t.FirstSupportMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = t.FirstSupportMan)) AS FirstSupportManName, t.SecondSupportMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = t.SecondSupportMan)) AS SecondSupportManName, t.Determination, t.FaultParts, t.FaultReason, t.Dispose, t.FollowUp, t.Improvements, t.Remark, t.BuDate, t.BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_4 WHERE (EMPNO = t.BuMan)) AS BuManName, t.ModifyDate, t.ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (EMPNO = t.ModifyMan)) AS ModifyManName, t.CarKMs, t.TowingCost, t.FixFee, t.LastServiceDate, t.DamageAnalysis, (SELECT CLASSTXT FROM DBDICB AS DBDICB_3 WHERE (FKEY = '拖吊車叫用記錄表TowTruckAssign  DamageAnalysis') AND (CLASSNO = t.DamageAnalysis)) AS DamageAnalysis_C, t.CaseType, (SELECT CLASSTXT FROM DBDICB AS DBDICB_2 WHERE (FKEY = '拖吊車叫用記錄表TowTruckAssign  CaseType') AND (CLASSNO = t.CaseType)) AS CaseType_C, t.WorkSheetNo, t.IsChecked FROM TowTruckAssignList AS t LEFT OUTER JOIN Car_infoA AS c ON c.Car_ID = t.Car_ID WHERE (t.CaseNo = @CaseNo)"
        DeleteCommand="DELETE FROM TowTruckAssignList WHERE (CaseNo = @CaseNo)"
        InsertCommand="INSERT INTO TowTruckAssignList(CaseNo, CaseDate, Car_ID, Driver, DepNo_Car, CasePosition, ParkingPosition, CaseTime, AssignMan, DepNo_Man, FirstSupportMan, SecondSupportMan, Determination, FaultParts, FaultReason, Dispose, FollowUp, Improvements, Remark, BuMan, BuDate, CarKMs, TowingCost, FixFee, LastServiceDate, DamageAnalysis, CaseType, WorkSheetNo, IsChecked) VALUES (@CaseNo, @CaseDate, @Car_ID, @Driver, @DepNo_Car, @CasePosition, @ParkingPosition, @CaseTime, @AssignMan, @DepNo_Man, @FirstSupportMan, @SecondSupportMan, @Determination, @FaultParts, @FaultReason, @Dispose, @FollowUp, @Improvements, @Remark, @BuMan, @BuDate, @CarKMs, @TowingCost, @FixFee, @LastServiceDate, @DamageAnalysis, @CaseType, @WorkSheetNo, @IsChecked)"
        UpdateCommand="UPDATE TowTruckAssignList SET CaseDate = @CaseDate, Car_ID = @Car_ID, Driver = @Driver, DepNo_Car = @DepNo_Car, CasePosition = @CasePosition, ParkingPosition = @ParkingPosition, CaseTime = @CaseTime, AssignMan = @AssignMan, DepNo_Man = @DepNo_Man, FirstSupportMan = @FirstSupportMan, SecondSupportMan = @SecondSupportMan, Determination = @Determination, FaultParts = @FaultParts, FaultReason = @FaultReason, Dispose = @Dispose, FollowUp = @FollowUp, Improvements = @Improvements, Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate, CarKMs = @CarKMs, TowingCost = @TowingCost, FixFee = @FixFee, LastServiceDate = @LastServiceDate, DamageAnalysis = @DamageAnalysis, CaseType = @CaseType, WorkSheetNo = @WorkSheetNo, IsChecked = @IsChecked WHERE (CaseNo = @CaseNo)"
        OnDeleted="sdsTowTruckAssignDetail_Deleted"
        OnInserted="sdsTowTruckAssignDetail_Inserted"
        OnUpdated="sdsTowTruckAssignDetail_Updated">
        <DeleteParameters>
            <asp:Parameter Name="CaseNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="CaseDate" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="DepNo_Car" />
            <asp:Parameter Name="CasePosition" />
            <asp:Parameter Name="ParkingPosition" />
            <asp:Parameter Name="CaseTime" />
            <asp:Parameter Name="AssignMan" />
            <asp:Parameter Name="DepNo_Man" />
            <asp:Parameter Name="FirstSupportMan" />
            <asp:Parameter Name="SecondSupportMan" />
            <asp:Parameter Name="Determination" />
            <asp:Parameter Name="FaultParts" />
            <asp:Parameter Name="FaultReason" />
            <asp:Parameter Name="Dispose" />
            <asp:Parameter Name="FollowUp" />
            <asp:Parameter Name="Improvements" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="BuDate" />
            <asp:Parameter Name="CarKMs" />
            <asp:Parameter Name="TowingCost" />
            <asp:Parameter Name="FixFee" />
            <asp:Parameter Name="LastServiceDate" />
            <asp:Parameter Name="DamageAnalysis" />
            <asp:Parameter Name="CaseType" />
            <asp:Parameter Name="WorkSheetNo" />
            <asp:Parameter Name="IsChecked" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridTowTruckAssignList" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="CaseDate" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="DepNo_Car" />
            <asp:Parameter Name="CasePosition" />
            <asp:Parameter Name="ParkingPosition" />
            <asp:Parameter Name="CaseTime" />
            <asp:Parameter Name="AssignMan" />
            <asp:Parameter Name="DepNo_Man" />
            <asp:Parameter Name="FirstSupportMan" />
            <asp:Parameter Name="SecondSupportMan" />
            <asp:Parameter Name="Determination" />
            <asp:Parameter Name="FaultParts" />
            <asp:Parameter Name="FaultReason" />
            <asp:Parameter Name="Dispose" />
            <asp:Parameter Name="FollowUp" />
            <asp:Parameter Name="Improvements" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="CarKMs" />
            <asp:Parameter Name="TowingCost" />
            <asp:Parameter Name="FixFee" />
            <asp:Parameter Name="LastServiceDate" />
            <asp:Parameter Name="DamageAnalysis" />
            <asp:Parameter Name="CaseType" />
            <asp:Parameter Name="WorkSheetNo" />
            <asp:Parameter Name="IsChecked" />
        </UpdateParameters>
    </asp:SqlDataSource>
</asp:Content>
