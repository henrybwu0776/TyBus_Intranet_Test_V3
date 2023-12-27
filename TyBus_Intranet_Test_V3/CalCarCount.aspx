<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CalCarCount.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CalCarCount" %>

<asp:Content ID="CalCarCountForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">各站車輛里程數計算</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:Label ID="lbCalYM_Search" runat="server" CssClass="text-Right-Blue" Text="計算年月(3 碼民國年 + 2 碼月份)" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eCalYM_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" Width="95%" OnClick="bbSearch_Click" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbCalData" runat="server" CssClass="button-Blue" Text="計算" Width="95%" OnClick="bbCalData_Click" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" Width="95%" OnClick="bbExit_Click" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plCarCount" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gvCarCount_List" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="DepNoYM" DataSourceID="dsCarCount_List" ForeColor="#333333" GridLines="None" PageSize="5">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="CalYM" DataFormatString="{0:d}" HeaderText="計算年月" SortExpression="CalYM" />
                <asp:BoundField DataField="DepNoYM" HeaderText="序號" ReadOnly="True" SortExpression="DepNoYM" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="站別" SortExpression="DepName" />
                <asp:BoundField DataField="UsedCars" HeaderText="使用車輛數" SortExpression="UsedCars" />
                <asp:BoundField DataField="CarCount" HeaderText="車輛總數" SortExpression="CarCount" />
                <asp:BoundField DataField="EmpCount" HeaderText="人員數" SortExpression="EmpCount" />
            </Columns>
            <EditRowStyle BackColor="#2461BF" />
            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#EFF3FB" />
            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#F5F7FB" />
            <SortedAscendingHeaderStyle BackColor="#6D95E1" />
            <SortedDescendingCellStyle BackColor="#E9EBEF" />
            <SortedDescendingHeaderStyle BackColor="#4870BE" />
        </asp:GridView>
        <asp:FormView ID="fvCarCount" runat="server" Width="100%" DataKeyNames="DepNoYM" DataSourceID="dsCarCount">
            <EditItemTemplate>                
                <asp:Button ID="bbEdit_OK" runat="server" CausesValidation="True" OnClick="UpdateButton_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel" runat="server" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_Edit_1" runat="server" CssClass="titleText-Red" Text="車輛數" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCalYM_Edit" runat="server" CssClass="text-Right-Blue" Text="統計年月" Width="95%" />
                                    <asp:Label ID="eDepNoYM_Edit" runat="server" Text='<%# Eval("DepNoYM") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCalYM_Edit" runat="server" Text='<%# Eval("CalYM") %>' CssClass="text-Left-Black" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="站別" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:Label ID="eDepNo_Edit" runat="server" Text='<%# Eval("DepNo") %>' CssClass="text-Left-Black" Width="30%" />
                                    <asp:Label ID="eDepName_Edit" runat="server" Text='<%# Eval("DepName") %>' CssClass="text-Left-Black" Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbUsedCars_Edit" runat="server" CssClass="text-Right-Blue" Text="車輛使用數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eUsedCars_Edit" runat="server" Text='<%# Eval("UsedCars") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbEmpCount_Edit" runat="server" CssClass="text-Right-Blue" Text="站員工數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eEmpCount_Edit" runat="server" Text='<%# Eval("EmpCount") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbEmpCount_M_Edit" runat="server" CssClass="text-Right-Blue" Text="管理人員人數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eEmpCount_M_Edit" runat="server" Text='<%# Eval("EmpCount_M") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eEmpCount_M_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbEmpCount_F_Edit" runat="server" CssClass="text-Right-Blue" Text="修護人員人數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eEmpCount_F_Edit" runat="server" Text='<%# Eval("EmpCount_F") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eEmpCount_M_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbEmpCount_D_Edit" runat="server" CssClass="text-Right-Blue" Text="行車人員人數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eEmpCount_D_Edit" runat="server" Text='<%# Eval("EmpCount_D") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eEmpCount_M_Edit_TextChanged" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecIncome_Edit" runat="server" CssClass="text-Right-Blue" Text="專車收入" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eSpecIncome_Edit" runat="server" Text='<%# Eval("SpecIncome") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="coolh ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbElecCardCount_Edit" runat="server" CssClass="text-Right-Blue" Text="電子票證載客人數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eElecCardCount_Edit" runat="server" Text='<%# Eval("ElecCardCount") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBreakDown_Edit" runat="server" CssClass="text-Right-Blue" Text="拋錨次數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBreakDown_Edit" runat="server" Text='<%# Eval("BreakDown") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarCount_Edit" runat="server" CssClass="text-Right-Blue" Text="站車輛總數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCarCount_Edit" runat="server" Text='<%# Eval("CarCount") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriverOD_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛加班時數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDriverOD_Edit" runat="server" Text='<%# Eval("DriverOD") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyYM_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyYM_Edit" runat="server" Text='<%# Eval("ModifyYM") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" Text='<%# Eval("ModifyMan") %>' CssClass="text-Left-Black" Width="30%" />
                                    <asp:Label ID="eEmpName_Edit" runat="server" Text='<%# Eval("EmpName") %>' CssClass="text-Left-Black" Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_Edit_2" runat="server" CssClass="titleText-S-Blue" Text="大巴車輛數" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleCount_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油料班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eRuleCount_OA_Edit" runat="server" Text='<%# Eval("RuleCount_OA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleCount_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電能班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eRuleCount_EA_Edit" runat="server" Text='<%# Eval("RuleCount_EA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbHighwayCount_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油料國道車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eHighWayCount_OA_Edit" runat="server" Text='<%# Eval("HighWayCount_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbHightwayCount_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電能國道車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eHighWayCount_EA_Edit" runat="server" Text='<%# Eval("HighWayCount_EA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusCount_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油料公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBusCount_OA_Edit" runat="server" Text='<%# Eval("BusCount_OA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusCount_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電能公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBusCount_EA_Edit" runat="server" Text='<%# Eval("BusCount_EA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourCount_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油料遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTourCount_OA_Edit" runat="server" Text='<%# Eval("TourCount_OA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourCount_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電能遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTourCount_EA_Edit" runat="server" Text='<%# Eval("TourCount_EA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbOtherCount_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油料其他" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eOtherCount_OA_Edit" runat="server" Text='<%# Eval("OtherCount_OA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbOtherCount_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電能其他" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eOtherCount_EA_Edit" runat="server" Text='<%# Eval("OtherCount_EA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSum_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油料車數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSum_OA_Edit" runat="server" Text='<%# Eval("CarSum_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSum_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電能車數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSum_EA_Edit" runat="server" Text='<%# Eval("CarSum_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSumA_Edit" runat="server" CssClass="text-Right-Blue" Text="大巴總數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSumA_Edit" runat="server" Text='<%# Eval("CarSumA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_Edit_3" runat="server" CssClass="titleText-S-Blue" Text="中巴車輛數" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleCount_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油料班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eRuleCount_OB_Edit" runat="server" Text='<%# Eval("RuleCount_OB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleCount_EB_Edit" runat="server" CssClass="text-Right-Blue" Text="電能班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eRuleCount_EB_Edit" runat="server" Text='<%# Eval("RuleCount_EB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbHighWayCount_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油料國道車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eHighWayCount_OB_Edit" runat="server" Text='<%# Eval("HighWayCount_OB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbHighWayCount_EB_Edit" runat="server" CssClass="text-Right-Blue" Text="電能國道車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eHighWayCount_EB_Edit" runat="server" Text='<%# Eval("HighWayCount_EB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusCount_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油料公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBusCount_OB_Edit" runat="server" Text='<%# Eval("BusCount_OB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusCount_EB_Edit" runat="server" CssClass="text-Right-Blue" Text="電能公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBusCount_EB_Edit" runat="server" Text='<%# Eval("BusCount_EB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourCount_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油料遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTourCount_OB_Edit" runat="server" Text='<%# Eval("TourCount_OB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourCount_EB_Edit" runat="server" CssClass="text-Right-Blue" Text="電能遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTourCount_EB_Edit" runat="server" Text='<%# Eval("TourCount_EB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbOtherCount_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油料其他" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eOtherCount_OB_Edit" runat="server" Text='<%# Eval("OtherCount_OB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbOtherCount_EB_Edit" runat="server" CssClass="text-Right-Blue" Text=" 電能其他" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eOtherCount_EB_Edit" runat="server" Text='<%# Eval("OtherCount_EB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleCount_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSum_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油料車數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSum_OB_Edit" runat="server" Text='<%# Eval("CarSum_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSum_EB_Edit" runat="server" CssClass="text-Right-Blue" Text="電能車數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSum_EB_Edit" runat="server" Text='<%# Eval("CarSum_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSumB_Edit" runat="server" CssClass="text-Right-Blue" Text="中巴總數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSumB_Edit" runat="server" Text='<%# Eval("CarSumB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_Edit_4" runat="server" CssClass="titleText-S-Blue" Text="站車輛數總計" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleCount_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleCount_OT_Edit" runat="server" Text='<%# Eval("RuleCount_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleCount_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電能班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleCount_ET_Edit" runat="server" Text='<%# Eval("RuleCount_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbHighWayCount_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料國道車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eHighWayCount_OT_Edit" runat="server" Text='<%# Eval("HighWayCount_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbHighWayCount_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電能國道車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eHighWayCount_ET_Edit" runat="server" Text='<%# Eval("HighWayCount_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusCount_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusCount_OT_Edit" runat="server" Text='<%# Eval("BusCount_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusCount_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電能公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusCount_ET_Edit" runat="server" Text='<%# Eval("BusCount_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourCount_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourCount_OT_Edit" runat="server" Text='<%# Eval("TourCount_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourCount_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電能遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourCount_ET_Edit" runat="server" Text='<%# Eval("TourCount_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbOtherCount_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料其他" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eOtherCount_OT_Edit" runat="server" Text='<%# Eval("OtherCount_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbOtherCount_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電能其他" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eOtherCount_ET_Edit" runat="server" Text='<%# Eval("OtherCount_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSum_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料車數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSum_OT_Edit" runat="server" Text='<%# Eval("CarSum_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSum_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電能車數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSum_ET_Edit" runat="server" Text='<%# Eval("CarSum_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_Edit_5" runat="server" CssClass="titleText-Red" Text="大巴里程數" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleKMS_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油料班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eRuleKMS_OA_Edit" runat="server" Text='<%# Eval("RuleKMS_OA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusKMS_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油料公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBusKMS_OA_Edit" runat="server" Text='<%# Eval("BusKMS_OA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecKMS_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油料專車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eSpecKMS_OA_Edit" runat="server" Text='<%# Eval("SpecKMS_OA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRentKMS_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油料租車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eRentKMS_OA_Edit" runat="server" Text='<%# Eval("RentKMS_OA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTransKMS_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油料交通車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTransKMS_OA_Edit" runat="server" Text='<%# Eval("TransKMS_OA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourKMS_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油料遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTourKMS_OA_Edit" runat="server" Text='<%# Eval("TourKMS_OA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNoneBusiKMS_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油料非營運" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eNoneBusiKMS_OA_Edit" runat="server" Text='<%# Eval("NoneBusiKMS_OA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTotalKMS_OA_Edit" runat="server" CssClass="text-Right-Blue" Text="油車里程小計" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTotalKMS_OA_Edit" runat="server" Text='<%# Eval("TotalKMS_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleKMS_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電能班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eRuleKMS_EA_Edit" runat="server" Text='<%# Eval("RuleKMS_EA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusKMS_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電能公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBusKMS_EA_Edit" runat="server" Text='<%# Eval("BusKMS_EA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecKMS_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電能專車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eSpecKMS_EA_Edit" runat="server" Text='<%# Eval("SpecKMS_EA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRentKMS_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電能租車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eRentKMS_EA_Edit" runat="server" Text='<%# Eval("RentKMS_EA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTransKMS_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電能交通車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTransKMS_EA_Edit" runat="server" Text='<%# Eval("TransKMS_EA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourKMS_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電能遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTourKMS_EA_Edit" runat="server" Text='<%# Eval("TourKMS_EA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNoneBusiKMS_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電能非營運" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eNoneBusiKMS_EA_Edit" runat="server" Text='<%# Eval("NoneBusiKMS_EA") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTotalKMS_EA_Edit" runat="server" CssClass="text-Right-Blue" Text="電車里程小計" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTotalKMS_EA_Edit" runat="server" Text='<%# Eval("TotalKMS_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_Edit_6" runat="server" CssClass="titleText-Red" Text="中巴里程數" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleKMS_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油料班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eRuleKMS_OB_Edit" runat="server" Text='<%# Eval("RuleKMS_OB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusKMS_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油料公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBusKMS_OB_Edit" runat="server" Text='<%# Eval("BusKMS_OB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecKMS_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油料專車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eSpecKMS_OB_Edit" runat="server" Text='<%# Eval("SpecKMS_OB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="coloh ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRentKMS_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油料租車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eRentKMS_OB_Edit" runat="server" Text='<%# Eval("RentKMS_OB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTransKMS_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油料交通車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTransKMS_OB_Edit" runat="server" Text='<%# Eval("TransKMS_OB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourKMS_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油料遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTourKMS_OB_Edit" runat="server" Text='<%# Eval("TourKMS_OB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNoneBusiKMS_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油料非營運" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eNoneBusiKMS_OB_Edit" runat="server" Text='<%# Eval("NoneBusiKMS_OB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTotalKMS_OB_Edit" runat="server" CssClass="text-Right-Blue" Text="油車里程小計" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTotalKMS_OB_Edit" runat="server" Text='<%# Eval("TotalKMS_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleKMS_EB_Edit" runat="server" CssClass="text-Right-Blue" Text="電能班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eRuleKMS_EB_Edit" runat="server" Text='<%# Eval("RuleKMS_EB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusKMS_EB_Edit" runat="server" CssClass="text-Right-Blue" Text="電能公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBusKMS_EB_Edit" runat="server" Text='<%# Eval("BusKMS_EB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecKMS_EB_Edit" runat="server" CssClass="text-Right-Blue" Text="電能專車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eSpecKMS_EB_Edit" runat="server" Text='<%# Eval("SpecKMS_EB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRentKMS_EB_Edit" runat="server" CssClass="text-Right-Blue" Text="電能租車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eRentKMS_EB_Edit" runat="server" Text='<%# Eval("RentKMS_EB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTransKMS_EB_Edit" runat="server" CssClass="text-Right-Blue" Text="電能交通車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTransKMS_EB_Edit" runat="server" Text='<%# Eval("TransKMS_EB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourKMS_EB_Edit" runat="server" CssClass="text-Right-Blue" Text="電能遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTourKMS_EB_Edit" runat="server" Text='<%# Eval("TourKMS_EB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNoneBusiKMS_EB_Edit" runat="server" CssClass="text-Right-Blue" Text="電能非營運" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eNoneBusiKMS_EB_Edit" runat="server" Text='<%# Eval("NoneBusiKMS_EB") %>' CssClass="text-Right-Black" Width="95%" AutoPostBack="true" OnTextChanged="eRuleKMS_OA_Edit_TextChanged" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTotalKMS_EB_Edit" runat="server" CssClass="text-Right-Blue" Text="電車里程小計" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTotalKMS_EB_Edit" runat="server" Text='<%# Eval("TotalKMS_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_Edit_7" runat="server" CssClass="titleText-S-Blue" Text="車種里程小計" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleKMS_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleKMS_OT_Edit" runat="server" Text='<%# Eval("RuleKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusKMS_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料公車" Width="95%" /> 
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusKMS_OT_Edit" runat="server" Text='<%# Eval("BusKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecKMS_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料專車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eSpecKMS_OT_Edit" runat="server" Text='<%# Eval("SpecKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRentKMS_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料租車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRentKMS_OT_Edit" runat="server" Text='<%# Eval("RentKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTransKMS_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料交通車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTransKMS_OT_Edit" runat="server" Text='<%# Eval("TransKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourKMS_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourKMS_OT_Edit" runat="server" Text='<%# Eval("TourKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNoneBusiKMS_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料非營運" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eNoneBusiKMS_OT_Edit" runat="server" Text='<%# Eval("NoneBusiKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTotalKMS_OT_Edit" runat="server" CssClass="text-Right-Blue" Text="油料車合計" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTotalKMS_OT_Edit" runat="server" Text='<%# Eval("TotalKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleKMS_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電能班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleKMS_ET_Edit" runat="server" Text='<%# Eval("RuleKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusKMS_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電能公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusKMS_ET_Edit" runat="server" Text='<%# Eval("BusKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecKMS_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電能專車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eSpecKMS_ET_Edit" runat="server" Text='<%# Eval("SpecKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRentKMS_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電能租車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRentKMS_ET_Edit" runat="server" Text='<%# Eval("RentKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTransKMS_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電能交通車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTransKMS_ET_Edit" runat="server" Text='<%# Eval("TransKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourKMS_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電能遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourKMS_ET_Edit" runat="server" Text='<%# Eval("TourKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNoneBusiKMS_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電車非營運" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eNoneBusiKMS_ET_Edit" runat="server" Text='<%# Eval("NoneBusiKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTotalKMS_ET_Edit" runat="server" CssClass="text-Right-Blue" Text="電能車小計" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTotalKMS_ET_Edit" runat="server" Text='<%# Eval("TotalKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_Edit_8" runat="server" CssClass="titleText-Red" Text="里程數合計" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_Rule_Edit" runat="server" CssClass="text-Right-Blue" Text="班車總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_Rule_Edit" runat="server" Text='<%# Eval("KMS_Rule") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_Bus_Edit" runat="server" CssClass="text-Right-Blue" Text="公車總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_BUS_Edit" runat="server" Text='<%# Eval("KMS_BUS") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_Spec_Edit" runat="server" CssClass="text-Right-Blue" Text="專車總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_Spec_Edit" runat="server" Text='<%# Eval("KMS_Spec") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_NoneBusi_Edit" runat="server" CssClass="text-Right-Blue" Text="非營運總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_NoneBusi_Edit" runat="server" Text='<%# Eval("KMS_NoneBusi") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_Rent_Edit" runat="server" CssClass="text-Right-Blue" Text="租車總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_Rent_Edit" runat="server" Text='<%# Eval("KMS_Rent") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_Trans_Edit" runat="server" CssClass="text-Right-Blue" Text="交通車總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_Trans_Edit" runat="server" Text='<%# Eval("KMS_Trans") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_Tour_Edit" runat="server" CssClass="text-Right-Blue" Text="遊覽車總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_Tour_Edit" runat="server" Text='<%# Eval("KMS_Tour") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriveRange_Edit" runat="server" CssClass="text-Right-Blue" Text="站別里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eDriveRange_Edit" runat="server" Text='<%# Eval("DriveRange") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>                            
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" style="visibility:hidden">
                                    <asp:Label ID="lbModifyKMS_Edit" runat="server" CssClass="text-Right-Red" Text="調整里程數" Width="95%" Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" style="visibility:hidden">
                                    <asp:TextBox ID="eModifyKMS_Edit" runat="server" Text='<%# Eval("ModifyKMS") %>' CssClass="text-Right-Black" Width="95%" Visible="false" />
                                    <asp:Label ID="eKMS_Total_Edit" runat="server" Text='<%# Eval("KMS_Total") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="titleText-S-Blue" Text="備註" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="8">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" Text='<%# Eval("Remark") %>' TextMode="MultiLine" CssClass="text-Left-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <InsertItemTemplate>
                
            </InsertItemTemplate>
            <ItemTemplate>
                <asp:Button ID="bbEdit_List" runat="server" CssClass="button-Blue" Text="更正" Width="120px" OnClick="bbEdit_List_Click" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_1" runat="server" CssClass="titleText-Red" Text="車輛數" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCalYM_List" runat="server" CssClass="text-Right-Blue" Text="統計年月" Width="95%" />
                                    <asp:Label ID="eDepNoYM_List" runat="server" Text='<%# Eval("DepNoYM") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCalYM_List" runat="server" Text='<%# Eval("CalYM") %>' CssClass="text-Left-Black" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="站別" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:Label ID="eDepNo_List" runat="server" Text='<%# Eval("DepNo") %>' CssClass="text-Left-Black" Width="30%" />
                                    <asp:Label ID="eDepName_List" runat="server" Text='<%# Eval("DepName") %>' CssClass="text-Left-Black" Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbUsedCars_List" runat="server" CssClass="text-Right-Blue" Text="車輛使用數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eUsedCars_List" runat="server" Text='<%# Eval("UsedCars") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbEmpCount_List" runat="server" CssClass="text-Right-Blue" Text="站員工數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eEmpCount_List" runat="server" Text='<%# Eval("EmpCount") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbEmpCount_M_List" runat="server" CssClass="text-Right-Blue" Text="管理人員人數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eEmpCount_M_List" runat="server" Text='<%# Eval("EmpCount_M") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbEmpCount_F_List" runat="server" CssClass="text-Right-Blue" Text="修護人員人數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eEmpCount_F_List" runat="server" Text='<%# Eval("EmpCount_F") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbEmpCount_D_List" runat="server" CssClass="text-Right-Blue" Text="行車人員人數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eEmpCount_D_List" runat="server" Text='<%# Eval("EmpCount_D") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecIncome_List" runat="server" CssClass="text-Right-Blue" Text="專車收入" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eSpecIncome_List" runat="server" Text='<%# Eval("SpecIncome") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="coolh ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbElecCardCount_List" runat="server" CssClass="text-Right-Blue" Text="電子票證載客人數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eElecCardCount_List" runat="server" Text='<%# Eval("ElecCardCount") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBreakDown_List" runat="server" CssClass="text-Right-Blue" Text="拋錨次數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBreakDown_List" runat="server" Text='<%# Eval("BreakDown") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarCount_List" runat="server" CssClass="text-Right-Blue" Text="站車輛總數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarCount_List" runat="server" Text='<%# Eval("CarCount") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriverOD_List" runat="server" CssClass="text-Right-Blue" Text="駕駛加班時數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eDriverOD_List" runat="server" Text='<%# Eval("DriverOD") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyYM_List" runat="server" CssClass="text-Right-Blue" Text="異動日期" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyYM_List" runat="server" Text='<%# Eval("ModifyYM") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:Label ID="eModifyMan_List" runat="server" Text='<%# Eval("ModifyMan") %>' CssClass="text-Left-Black" Width="30%" />
                                    <asp:Label ID="eEmpName_List" runat="server" Text='<%# Eval("EmpName") %>' CssClass="text-Left-Black" Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_2" runat="server" CssClass="titleText-S-Blue" Text="大巴車輛數" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleCount_OA_List" runat="server" CssClass="text-Right-Blue" Text="油料班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleCount_OA_List" runat="server" Text='<%# Eval("RuleCount_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleCount_EA_List" runat="server" CssClass="text-Right-Blue" Text="電能班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleCount_EA_List" runat="server" Text='<%# Eval("RuleCount_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbHighwayCount_OA_List" runat="server" CssClass="text-Right-Blue" Text="油料國道車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eHighWayCount_OA_List" runat="server" Text='<%# Eval("HighWayCount_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbHightwayCount_EA_List" runat="server" CssClass="text-Right-Blue" Text="電能國道車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eHighWayCount_EA_List" runat="server" Text='<%# Eval("HighWayCount_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusCount_OA_List" runat="server" CssClass="text-Right-Blue" Text="油料公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusCount_OA_List" runat="server" Text='<%# Eval("BusCount_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusCount_EA_List" runat="server" CssClass="text-Right-Blue" Text="電能公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusCount_EA_List" runat="server" Text='<%# Eval("BusCount_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourCount_OA_List" runat="server" CssClass="text-Right-Blue" Text="油料遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourCount_OA_List" runat="server" Text='<%# Eval("TourCount_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourCount_EA_List" runat="server" CssClass="text-Right-Blue" Text="電能遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourCount_EA_List" runat="server" Text='<%# Eval("TourCount_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbOtherCount_OA_List" runat="server" CssClass="text-Right-Blue" Text="油料其他" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eOtherCount_OA_List" runat="server" Text='<%# Eval("OtherCount_OA") %>' CssClass="text-Right-Black" Width="95%     " />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbOtherCount_EA_List" runat="server" CssClass="text-Right-Blue" Text="電能其他" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eOtherCount_EA_List" runat="server" Text='<%# Eval("OtherCount_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSum_OA_List" runat="server" CssClass="text-Right-Blue" Text="油料車數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSum_OA_List" runat="server" Text='<%# Eval("CarSum_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSum_EA_List" runat="server" CssClass="text-Right-Blue" Text="電能車數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSum_EA_List" runat="server" Text='<%# Eval("CarSum_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSumA_List" runat="server" CssClass="text-Right-Blue" Text="大巴總數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSumA_List" runat="server" Text='<%# Eval("CarSumA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_3" runat="server" CssClass="titleText-S-Blue" Text="中巴車輛數" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleCount_OB_List" runat="server" CssClass="text-Right-Blue" Text="油料班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleCount_OB_List" runat="server" Text='<%# Eval("RuleCount_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleCount_EB_List" runat="server" CssClass="text-Right-Blue" Text="電能班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleCount_EB_List" runat="server" Text='<%# Eval("RuleCount_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbHighWayCount_OB_List" runat="server" CssClass="text-Right-Blue" Text="油料國道車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eHighWayCount_OB_List" runat="server" Text='<%# Eval("HighWayCount_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbHighWayCount_EB_List" runat="server" CssClass="text-Right-Blue" Text="電能國道車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eHighWayCount_EB_List" runat="server" Text='<%# Eval("HighWayCount_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusCount_OB_List" runat="server" CssClass="text-Right-Blue" Text="油料公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusCount_OB_List" runat="server" Text='<%# Eval("BusCount_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusCount_EB_List" runat="server" CssClass="text-Right-Blue" Text="電能公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusCount_EB_List" runat="server" Text='<%# Eval("BusCount_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourCount_OB_List" runat="server" CssClass="text-Right-Blue" Text="油料遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourCount_OB_List" runat="server" Text='<%# Eval("TourCount_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourCount_EB_List" runat="server" CssClass="text-Right-Blue" Text="電能遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourCount_EB_List" runat="server" Text='<%# Eval("TourCount_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbOtherCount_OB_List" runat="server" CssClass="text-Right-Blue" Text="油料其他" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eOtherCount_OB_List" runat="server" Text='<%# Eval("OtherCount_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbOtherCount_EB_List" runat="server" CssClass="text-Right-Blue" Text=" 電能其他" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eOtherCount_EB_List" runat="server" Text='<%# Eval("OtherCount_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSum_OB_List" runat="server" CssClass="text-Right-Blue" Text="油料車數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSum_OB_List" runat="server" Text='<%# Eval("CarSum_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSum_EB_List" runat="server" CssClass="text-Right-Blue" Text="電能車數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSum_EB_List" runat="server" Text='<%# Eval("CarSum_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSumB_List" runat="server" CssClass="text-Right-Blue" Text="中巴總數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSumB_List" runat="server" Text='<%# Eval("CarSumB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_4" runat="server" CssClass="titleText-S-Blue" Text="站車輛數總計" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleCount_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleCount_OT_List" runat="server" Text='<%# Eval("RuleCount_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleCount_ET_List" runat="server" CssClass="text-Right-Blue" Text="電能班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleCount_ET_List" runat="server" Text='<%# Eval("RuleCount_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbHighWayCount_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料國道車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eHighWayCount_OT_List" runat="server" Text='<%# Eval("HighWayCount_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbHighWayCount_ET_List" runat="server" CssClass="text-Right-Blue" Text="電能國道車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eHighWayCount_ET_List" runat="server" Text='<%# Eval("HighWayCount_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusCount_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusCount_OT_List" runat="server" Text='<%# Eval("BusCount_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusCount_ET_List" runat="server" CssClass="text-Right-Blue" Text="電能公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusCount_ET_List" runat="server" Text='<%# Eval("BusCount_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourCount_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourCount_OT_List" runat="server" Text='<%# Eval("TourCount_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourCount_ET_List" runat="server" CssClass="text-Right-Blue" Text="電能遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourCount_ET_List" runat="server" Text='<%# Eval("TourCount_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbOtherCount_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料其他" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eOtherCount_OT_List" runat="server" Text='<%# Eval("OtherCount_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbOtherCount_ET_List" runat="server" CssClass="text-Right-Blue" Text="電能其他" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eOtherCount_ET_List" runat="server" Text='<%# Eval("OtherCount_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSum_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料車數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSum_OT_List" runat="server" Text='<%# Eval("CarSum_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarSum_ET_List" runat="server" CssClass="text-Right-Blue" Text="電能車數" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarSum_ET_List" runat="server" Text='<%# Eval("CarSum_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_5" runat="server" CssClass="titleText-Red" Text="大巴里程數" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleKMS_OA_List" runat="server" CssClass="text-Right-Blue" Text="油料班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleKMS_OA_List" runat="server" Text='<%# Eval("RuleKMS_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusKMS_OA_List" runat="server" CssClass="text-Right-Blue" Text="油料公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusKMS_OA_List" runat="server" Text='<%# Eval("BusKMS_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecKMS_OA_List" runat="server" CssClass="text-Right-Blue" Text="油料專車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eSpecKMS_OA_List" runat="server" Text='<%# Eval("SpecKMS_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRentKMS_OA_List" runat="server" CssClass="text-Right-Blue" Text="油料租車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRentKMS_OA_List" runat="server" Text='<%# Eval("RentKMS_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTransKMS_OA_List" runat="server" CssClass="text-Right-Blue" Text="油料交通車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTransKMS_OA_List" runat="server" Text='<%# Eval("TransKMS_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourKMS_OA_List" runat="server" CssClass="text-Right-Blue" Text="油料遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourKMS_OA_List" runat="server" Text='<%# Eval("TourKMS_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNoneBusiKMS_OA_List" runat="server" CssClass="text-Right-Blue" Text="油料非營運" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eNoneBusiKMS_OA_List" runat="server" Text='<%# Eval("NoneBusiKMS_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTotalKMS_OA_List" runat="server" CssClass="text-Right-Blue" Text="油車里程小計" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTotalKMS_OA_List" runat="server" Text='<%# Eval("TotalKMS_OA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleKMS_EA_List" runat="server" CssClass="text-Right-Blue" Text="電能班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleKMS_EA_List" runat="server" Text='<%# Eval("RuleKMS_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusKMS_EA_List" runat="server" CssClass="text-Right-Blue" Text="電能公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusKMS_EA_List" runat="server" Text='<%# Eval("BusKMS_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecKMS_EA_List" runat="server" CssClass="text-Right-Blue" Text="電能專車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eSpecKMS_EA_List" runat="server" Text='<%# Eval("SpecKMS_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRentKMS_EA_List" runat="server" CssClass="text-Right-Blue" Text="電能租車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRentKMS_EA_List" runat="server" Text='<%# Eval("RentKMS_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTransKMS_EA_List" runat="server" CssClass="text-Right-Blue" Text="電能交通車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTransKMS_EA_List" runat="server" Text='<%# Eval("TransKMS_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourKMS_EA_List" runat="server" CssClass="text-Right-Blue" Text="電能遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourKMS_EA_List" runat="server" Text='<%# Eval("TourKMS_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNoneBusiKMS_EA_List" runat="server" CssClass="text-Right-Blue" Text="電能非營運" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eNoneBusiKMS_EA_List" runat="server" Text='<%# Eval("NoneBusiKMS_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTotalKMS_EA_List" runat="server" CssClass="text-Right-Blue" Text="電車里程小計" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTotalKMS_EA_List" runat="server" Text='<%# Eval("TotalKMS_EA") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_6" runat="server" CssClass="titleText-Red" Text="中巴里程數" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleKMS_OB_List" runat="server" CssClass="text-Right-Blue" Text="油料班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleKMS_OB_List" runat="server" Text='<%# Eval("RuleKMS_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusKMS_OB_List" runat="server" CssClass="text-Right-Blue" Text="油料公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusKMS_OB_List" runat="server" Text='<%# Eval("BusKMS_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecKMS_OB_List" runat="server" CssClass="text-Right-Blue" Text="油料專車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eSpecKMS_OB_List" runat="server" Text='<%# Eval("SpecKMS_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="coloh ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRentKMS_OB_List" runat="server" CssClass="text-Right-Blue" Text="油料租車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRentKMS_OB_List" runat="server" Text='<%# Eval("RentKMS_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTransKMS_OB_List" runat="server" CssClass="text-Right-Blue" Text="油料交通車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTransKMS_OB_List" runat="server" Text='<%# Eval("TransKMS_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourKMS_OB_List" runat="server" CssClass="text-Right-Blue" Text="油料遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourKMS_OB_List" runat="server" Text='<%# Eval("TourKMS_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNoneBusiKMS_OB_List" runat="server" CssClass="text-Right-Blue" Text="油料非營運" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eNoneBusiKMS_OB_List" runat="server" Text='<%# Eval("NoneBusiKMS_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTotalKMS_OB_List" runat="server" CssClass="text-Right-Blue" Text="油車里程小計" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTotalKMS_OB_List" runat="server" Text='<%# Eval("TotalKMS_OB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleKMS_EB_List" runat="server" CssClass="text-Right-Blue" Text="電能班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleKMS_EB_List" runat="server" Text='<%# Eval("RuleKMS_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusKMS_EB_List" runat="server" CssClass="text-Right-Blue" Text="電能公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusKMS_EB_List" runat="server" Text='<%# Eval("BusKMS_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecKMS_EB_List" runat="server" CssClass="text-Right-Blue" Text="電能專車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eSpecKMS_EB_List" runat="server" Text='<%# Eval("SpecKMS_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRentKMS_EB_List" runat="server" CssClass="text-Right-Blue" Text="電能租車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRentKMS_EB_List" runat="server" Text='<%# Eval("RentKMS_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTransKMS_EB_List" runat="server" CssClass="text-Right-Blue" Text="電能交通車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTransKMS_EB_List" runat="server" Text='<%# Eval("TransKMS_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourKMS_EB_List" runat="server" CssClass="text-Right-Blue" Text="電能遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourKMS_EB_List" runat="server" Text='<%# Eval("TourKMS_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNoneBusiKMS_EB_List" runat="server" CssClass="text-Right-Blue" Text="電能非營運" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eNoneBusiKMS_EB_List" runat="server" Text='<%# Eval("NoneBusiKMS_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTotalKMS_EB_List" runat="server" CssClass="text-Right-Blue" Text="電車里程小計" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTotalKMS_EB_List" runat="server" Text='<%# Eval("TotalKMS_EB") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_7" runat="server" CssClass="titleText-S-Blue" Text="車種里程小計" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleKMS_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleKMS_OT_List" runat="server" Text='<%# Eval("RuleKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusKMS_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料公車" Width="95%" /> 
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusKMS_OT_List" runat="server" Text='<%# Eval("BusKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecKMS_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料專車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eSpecKMS_OT_List" runat="server" Text='<%# Eval("SpecKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRentKMS_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料租車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRentKMS_OT_List" runat="server" Text='<%# Eval("RentKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTransKMS_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料交通車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTransKMS_OT_List" runat="server" Text='<%# Eval("TransKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourKMS_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourKMS_OT_List" runat="server" Text='<%# Eval("TourKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNoneBusiKMS_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料非營運" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eNoneBusiKMS_OT_List" runat="server" Text='<%# Eval("NoneBusiKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTotalKMS_OT_List" runat="server" CssClass="text-Right-Blue" Text="油料車合計" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTotalKMS_OT_List" runat="server" Text='<%# Eval("TotalKMS_OT") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRuleKMS_ET_List" runat="server" CssClass="text-Right-Blue" Text="電能班車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRuleKMS_ET_List" runat="server" Text='<%# Eval("RuleKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBusKMS_ET_List" runat="server" CssClass="text-Right-Blue" Text="電能公車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBusKMS_ET_List" runat="server" Text='<%# Eval("BusKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSpecKMS_ET_List" runat="server" CssClass="text-Right-Blue" Text="電能專車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eSpecKMS_ET_List" runat="server" Text='<%# Eval("SpecKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRentKMS_ET_List" runat="server" CssClass="text-Right-Blue" Text="電能租車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eRentKMS_ET_List" runat="server" Text='<%# Eval("RentKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTransKMS_ET_List" runat="server" CssClass="text-Right-Blue" Text="電能交通車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTransKMS_ET_List" runat="server" Text='<%# Eval("TransKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTourKMS_ET_List" runat="server" CssClass="text-Right-Blue" Text="電能遊覽車" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTourKMS_ET_List" runat="server" Text='<%# Eval("TourKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNoneBusiKMS_ET_List" runat="server" CssClass="text-Right-Blue" Text="電車非營運" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eNoneBusiKMS_ET_List" runat="server" Text='<%# Eval("NoneBusiKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbTotalKMS_ET_List" runat="server" CssClass="text-Right-Blue" Text="電能車小計" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTotalKMS_ET_List" runat="server" Text='<%# Eval("TotalKMS_ET") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_8" runat="server" CssClass="titleText-Red" Text="里程數合計" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_Rule_List" runat="server" CssClass="text-Right-Blue" Text="班車總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_Rule_List" runat="server" Text='<%# Eval("KMS_Rule") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_Bus_List" runat="server" CssClass="text-Right-Blue" Text="公車總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_BUS_List" runat="server" Text='<%# Eval("KMS_BUS") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_Spec_List" runat="server" CssClass="text-Right-Blue" Text="專車總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_Spec_List" runat="server" Text='<%# Eval("KMS_Spec") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_NoneBusi_List" runat="server" CssClass="text-Right-Blue" Text="非營運總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_NoneBusi_List" runat="server" Text='<%# Eval("KMS_NoneBusi") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_Rent_List" runat="server" CssClass="text-Right-Blue" Text="租車總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_Rent_List" runat="server" Text='<%# Eval("KMS_Rent") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_Trans_List" runat="server" CssClass="text-Right-Blue" Text="交通車總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_Trans_List" runat="server" Text='<%# Eval("KMS_Trans") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbKMS_Tour_List" runat="server" CssClass="text-Right-Blue" Text="遊覽車總里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eKMS_Tour_List" runat="server" Text='<%# Eval("KMS_Tour") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriveRange_List" runat="server" CssClass="text-Right-Blue" Text="站別里程" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eDriveRange_List" runat="server" Text='<%# Eval("DriveRange") %>' CssClass="text-Right-Black" Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" style="visibility:hidden">
                                    <asp:Label ID="lbModifyKMS_List" runat="server" CssClass="text-Right-Red" Text="調整里程數" Width="95%" Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" style="visibility:hidden">
                                    <asp:Label ID="eModifyKMS_List" runat="server" Text='<%# Eval("ModifyKMS") %>' CssClass="text-Right-Black" Width="95%" Visible="false" />
                                    <asp:Label ID="eKMS_Total_List" runat="server" Text='<%# Eval("KMS_Total") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbRemark_List" runat="server" CssClass="titleText-S-Blue" Text="備註" Width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="8">
                                    <asp:TextBox ID="eRemark_List" runat="server" Text='<%# Eval("Remark") %>' TextMode="MultiLine" CssClass="text-Left-Black" Width="95%" Enabled="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColWidth-8Col" />
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="dsCarCount" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT DepNoYM, CalYM, c.DepNo, d.[Name] DepName, ModifyYM, ModifyMan, e.[Name] EmpName, CarCount, DriveRange, c.Remark, 
       KMS_BUS, KMS_Rent, KMS_Tour, KMS_Trans, KMS_Spec, RuleCount_OA, RuleCount_EA, RuleCount_OB, RuleCount_EB, RuleCount_OT, RuleCount_ET, 
	   BusCount_OA, BusCount_EA, BusCount_OB, BusCount_EB, BusCount_OT, BusCount_ET, TourCount_OA, TourCount_EA, TourCount_OB, TourCount_EB, TourCount_OT, TourCount_ET, 
	   CarSum_OA, CarSum_EA, CarSum_OB, CarSum_EB, CarSum_OT, CarSum_ET, RuleKMS_OA, RuleKMS_EA, RuleKMS_OB, RuleKMS_EB, RuleKMS_OT, RuleKMS_ET, KMS_Rule, 
	   BusKMS_OA, BusKMS_EA, BusKMS_OB, BusKMS_EB, BusKMS_OT, BusKMS_ET, SpecKMS_OA, SpecKMS_EA, SpecKMS_OB, SpecKMS_EB, SpecKMS_OT, SpecKMS_ET, 
	   RentKMS_OA, RentKMS_EA, RentKMS_OB, RentKMS_EB, RentKMS_OT, RentKMS_ET, TransKMS_OA, TransKMS_EA, TransKMS_OB, TransKMS_EB, TransKMS_OT, TransKMS_ET, 
	   TourKMS_OA, TourKMS_EA, TourKMS_OB, TourKMS_EB, TourKMS_OT, TourKMS_ET, NoneBusiKMS_OA, NoneBusiKMS_EA, NoneBusiKMS_OB, NoneBusiKMS_EB, NoneBusiKMS_OT,NoneBusiKMS_ET, KMS_NoneBusi, 
	   TotalKMS_OA, TotalKMS_EA, TotalKMS_OB, TotalKMS_EB, TotalKMS_OT, TotalKMS_ET, KMS_Total, HighWayCount_OA, HighWayCount_EA, HighWayCount_OB, HighWayCount_EB, HighWayCount_OT, HighWayCount_ET, 
	   OtherCount_OA, OtherCount_EA, OtherCount_OB, OtherCount_EB, OtherCount_OT, OtherCount_ET, ModifyKMS, SpecIncome, UsedCars, ElecCardCount, EmpCount, DriverOD, EmpCount_M, EmpCount_F, EmpCount_D, BreakDown,
       (CarSum_OA + CarSum_EA) CarSumA, (CarSum_OB + CarSum_EB) CarSumB
  FROM dbo.CarCount c left join Department d on d.DepNo = c.DepNo
                      left join Employee e on e.EmpNo = c.ModifyMan
 WHERE isnull(DepNoYM, '') = @DepNoYM">
        <SelectParameters>
            <asp:ControlParameter ControlID="gvCarCount_List" Name="DepNoYM" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsCarCount_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select c.CalYM, c.DepNoYM, c.DepNo, d.[Name] DepName, UsedCars, CarCount, EmpCount
  from CarCount c left join Department d on d.DepNo = c.DepNo
 where isnull(DepNoYM, '') = ''"></asp:SqlDataSource>
</asp:Content>
