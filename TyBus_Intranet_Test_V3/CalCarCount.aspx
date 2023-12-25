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
        <asp:FormView ID="fvCarCount" runat="server" Width="100%" DataKeyNames="DepNoYM" DataSourceID="sdCarCount">
            <EditItemTemplate>
                DepNoYM:
                <asp:Label ID="DepNoYMLabel1" runat="server" Text='<%# Eval("DepNoYM") %>' />
                <br />
                CalYM:
                <asp:TextBox ID="CalYMTextBox" runat="server" Text='<%# Bind("CalYM") %>' />
                <br />
                DepNo:
                <asp:TextBox ID="DepNoTextBox" runat="server" Text='<%# Bind("DepNo") %>' />
                <br />
                depName:
                <asp:TextBox ID="depNameTextBox" runat="server" Text='<%# Bind("depName") %>' />
                <br />
                ModifyYM:
                <asp:TextBox ID="ModifyYMTextBox" runat="server" Text='<%# Bind("ModifyYM") %>' />
                <br />
                ModifyMan:
                <asp:TextBox ID="ModifyManTextBox" runat="server" Text='<%# Bind("ModifyMan") %>' />
                <br />
                CarCount:
                <asp:TextBox ID="CarCountTextBox" runat="server" Text='<%# Bind("CarCount") %>' />
                <br />
                DriveRange:
                <asp:TextBox ID="DriveRangeTextBox" runat="server" Text='<%# Bind("DriveRange") %>' />
                <br />
                Remark:
                <asp:TextBox ID="RemarkTextBox" runat="server" Text='<%# Bind("Remark") %>' />
                <br />
                KMS_BUS:
                <asp:TextBox ID="KMS_BUSTextBox" runat="server" Text='<%# Bind("KMS_BUS") %>' />
                <br />
                KMS_Rent:
                <asp:TextBox ID="KMS_RentTextBox" runat="server" Text='<%# Bind("KMS_Rent") %>' />
                <br />
                KMS_Tour:
                <asp:TextBox ID="KMS_TourTextBox" runat="server" Text='<%# Bind("KMS_Tour") %>' />
                <br />
                KMS_Trans:
                <asp:TextBox ID="KMS_TransTextBox" runat="server" Text='<%# Bind("KMS_Trans") %>' />
                <br />
                KMS_Spec:
                <asp:TextBox ID="KMS_SpecTextBox" runat="server" Text='<%# Bind("KMS_Spec") %>' />
                <br />
                RuleCount_OA:
                <asp:TextBox ID="RuleCount_OATextBox" runat="server" Text='<%# Bind("RuleCount_OA") %>' />
                <br />
                RuleCount_EA:
                <asp:TextBox ID="RuleCount_EATextBox" runat="server" Text='<%# Bind("RuleCount_EA") %>' />
                <br />
                RuleCount_OB:
                <asp:TextBox ID="RuleCount_OBTextBox" runat="server" Text='<%# Bind("RuleCount_OB") %>' />
                <br />
                RuleCount_EB:
                <asp:TextBox ID="RuleCount_EBTextBox" runat="server" Text='<%# Bind("RuleCount_EB") %>' />
                <br />
                RuleCount_OT:
                <asp:TextBox ID="RuleCount_OTTextBox" runat="server" Text='<%# Bind("RuleCount_OT") %>' />
                <br />
                RuleCount_ET:
                <asp:TextBox ID="RuleCount_ETTextBox" runat="server" Text='<%# Bind("RuleCount_ET") %>' />
                <br />
                BusCount_OA:
                <asp:TextBox ID="BusCount_OATextBox" runat="server" Text='<%# Bind("BusCount_OA") %>' />
                <br />
                BusCount_EA:
                <asp:TextBox ID="BusCount_EATextBox" runat="server" Text='<%# Bind("BusCount_EA") %>' />
                <br />
                BusCount_OB:
                <asp:TextBox ID="BusCount_OBTextBox" runat="server" Text='<%# Bind("BusCount_OB") %>' />
                <br />
                BusCount_EB:
                <asp:TextBox ID="BusCount_EBTextBox" runat="server" Text='<%# Bind("BusCount_EB") %>' />
                <br />
                BusCount_OT:
                <asp:TextBox ID="BusCount_OTTextBox" runat="server" Text='<%# Bind("BusCount_OT") %>' />
                <br />
                BusCount_ET:
                <asp:TextBox ID="BusCount_ETTextBox" runat="server" Text='<%# Bind("BusCount_ET") %>' />
                <br />
                TourCount_OA:
                <asp:TextBox ID="TourCount_OATextBox" runat="server" Text='<%# Bind("TourCount_OA") %>' />
                <br />
                TourCount_EA:
                <asp:TextBox ID="TourCount_EATextBox" runat="server" Text='<%# Bind("TourCount_EA") %>' />
                <br />
                TourCount_OB:
                <asp:TextBox ID="TourCount_OBTextBox" runat="server" Text='<%# Bind("TourCount_OB") %>' />
                <br />
                TourCount_EB:
                <asp:TextBox ID="TourCount_EBTextBox" runat="server" Text='<%# Bind("TourCount_EB") %>' />
                <br />
                TourCount_OT:
                <asp:TextBox ID="TourCount_OTTextBox" runat="server" Text='<%# Bind("TourCount_OT") %>' />
                <br />
                TourCount_ET:
                <asp:TextBox ID="TourCount_ETTextBox" runat="server" Text='<%# Bind("TourCount_ET") %>' />
                <br />
                CarSum_OA:
                <asp:TextBox ID="CarSum_OATextBox" runat="server" Text='<%# Bind("CarSum_OA") %>' />
                <br />
                CarSum_EA:
                <asp:TextBox ID="CarSum_EATextBox" runat="server" Text='<%# Bind("CarSum_EA") %>' />
                <br />
                CarSum_OB:
                <asp:TextBox ID="CarSum_OBTextBox" runat="server" Text='<%# Bind("CarSum_OB") %>' />
                <br />
                CarSum_EB:
                <asp:TextBox ID="CarSum_EBTextBox" runat="server" Text='<%# Bind("CarSum_EB") %>' />
                <br />
                CarSum_OT:
                <asp:TextBox ID="CarSum_OTTextBox" runat="server" Text='<%# Bind("CarSum_OT") %>' />
                <br />
                CarSum_ET:
                <asp:TextBox ID="CarSum_ETTextBox" runat="server" Text='<%# Bind("CarSum_ET") %>' />
                <br />
                RuleKMS_OA:
                <asp:TextBox ID="RuleKMS_OATextBox" runat="server" Text='<%# Bind("RuleKMS_OA") %>' />
                <br />
                RuleKMS_EA:
                <asp:TextBox ID="RuleKMS_EATextBox" runat="server" Text='<%# Bind("RuleKMS_EA") %>' />
                <br />
                RuleKMS_OB:
                <asp:TextBox ID="RuleKMS_OBTextBox" runat="server" Text='<%# Bind("RuleKMS_OB") %>' />
                <br />
                RuleKMS_EB:
                <asp:TextBox ID="RuleKMS_EBTextBox" runat="server" Text='<%# Bind("RuleKMS_EB") %>' />
                <br />
                RuleKMS_OT:
                <asp:TextBox ID="RuleKMS_OTTextBox" runat="server" Text='<%# Bind("RuleKMS_OT") %>' />
                <br />
                RuleKMS_ET:
                <asp:TextBox ID="RuleKMS_ETTextBox" runat="server" Text='<%# Bind("RuleKMS_ET") %>' />
                <br />
                KMS_Rule:
                <asp:TextBox ID="KMS_RuleTextBox" runat="server" Text='<%# Bind("KMS_Rule") %>' />
                <br />
                BusKMS_OA:
                <asp:TextBox ID="BusKMS_OATextBox" runat="server" Text='<%# Bind("BusKMS_OA") %>' />
                <br />
                BusKMS_EA:
                <asp:TextBox ID="BusKMS_EATextBox" runat="server" Text='<%# Bind("BusKMS_EA") %>' />
                <br />
                BusKMS_OB:
                <asp:TextBox ID="BusKMS_OBTextBox" runat="server" Text='<%# Bind("BusKMS_OB") %>' />
                <br />
                BusKMS_EB:
                <asp:TextBox ID="BusKMS_EBTextBox" runat="server" Text='<%# Bind("BusKMS_EB") %>' />
                <br />
                BusKMS_OT:
                <asp:TextBox ID="BusKMS_OTTextBox" runat="server" Text='<%# Bind("BusKMS_OT") %>' />
                <br />
                BusKMS_ET:
                <asp:TextBox ID="BusKMS_ETTextBox" runat="server" Text='<%# Bind("BusKMS_ET") %>' />
                <br />
                SpecKMS_OA:
                <asp:TextBox ID="SpecKMS_OATextBox" runat="server" Text='<%# Bind("SpecKMS_OA") %>' />
                <br />
                SpecKMS_EA:
                <asp:TextBox ID="SpecKMS_EATextBox" runat="server" Text='<%# Bind("SpecKMS_EA") %>' />
                <br />
                SpecKMS_OB:
                <asp:TextBox ID="SpecKMS_OBTextBox" runat="server" Text='<%# Bind("SpecKMS_OB") %>' />
                <br />
                SpecKMS_EB:
                <asp:TextBox ID="SpecKMS_EBTextBox" runat="server" Text='<%# Bind("SpecKMS_EB") %>' />
                <br />
                SpecKMS_OT:
                <asp:TextBox ID="SpecKMS_OTTextBox" runat="server" Text='<%# Bind("SpecKMS_OT") %>' />
                <br />
                SpecKMS_ET:
                <asp:TextBox ID="SpecKMS_ETTextBox" runat="server" Text='<%# Bind("SpecKMS_ET") %>' />
                <br />
                RentKMS_OA:
                <asp:TextBox ID="RentKMS_OATextBox" runat="server" Text='<%# Bind("RentKMS_OA") %>' />
                <br />
                RentKMS_EA:
                <asp:TextBox ID="RentKMS_EATextBox" runat="server" Text='<%# Bind("RentKMS_EA") %>' />
                <br />
                RentKMS_OB:
                <asp:TextBox ID="RentKMS_OBTextBox" runat="server" Text='<%# Bind("RentKMS_OB") %>' />
                <br />
                RentKMS_EB:
                <asp:TextBox ID="RentKMS_EBTextBox" runat="server" Text='<%# Bind("RentKMS_EB") %>' />
                <br />
                RentKMS_OT:
                <asp:TextBox ID="RentKMS_OTTextBox" runat="server" Text='<%# Bind("RentKMS_OT") %>' />
                <br />
                RentKMS_ET:
                <asp:TextBox ID="RentKMS_ETTextBox" runat="server" Text='<%# Bind("RentKMS_ET") %>' />
                <br />
                TransKMS_OA:
                <asp:TextBox ID="TransKMS_OATextBox" runat="server" Text='<%# Bind("TransKMS_OA") %>' />
                <br />
                TransKMS_EA:
                <asp:TextBox ID="TransKMS_EATextBox" runat="server" Text='<%# Bind("TransKMS_EA") %>' />
                <br />
                TransKMS_OB:
                <asp:TextBox ID="TransKMS_OBTextBox" runat="server" Text='<%# Bind("TransKMS_OB") %>' />
                <br />
                TransKMS_EB:
                <asp:TextBox ID="TransKMS_EBTextBox" runat="server" Text='<%# Bind("TransKMS_EB") %>' />
                <br />
                TransKMS_OT:
                <asp:TextBox ID="TransKMS_OTTextBox" runat="server" Text='<%# Bind("TransKMS_OT") %>' />
                <br />
                TransKMS_ET:
                <asp:TextBox ID="TransKMS_ETTextBox" runat="server" Text='<%# Bind("TransKMS_ET") %>' />
                <br />
                TourKMS_OA:
                <asp:TextBox ID="TourKMS_OATextBox" runat="server" Text='<%# Bind("TourKMS_OA") %>' />
                <br />
                TourKMS_EA:
                <asp:TextBox ID="TourKMS_EATextBox" runat="server" Text='<%# Bind("TourKMS_EA") %>' />
                <br />
                TourKMS_OB:
                <asp:TextBox ID="TourKMS_OBTextBox" runat="server" Text='<%# Bind("TourKMS_OB") %>' />
                <br />
                TourKMS_EB:
                <asp:TextBox ID="TourKMS_EBTextBox" runat="server" Text='<%# Bind("TourKMS_EB") %>' />
                <br />
                TourKMS_OT:
                <asp:TextBox ID="TourKMS_OTTextBox" runat="server" Text='<%# Bind("TourKMS_OT") %>' />
                <br />
                TourKMS_ET:
                <asp:TextBox ID="TourKMS_ETTextBox" runat="server" Text='<%# Bind("TourKMS_ET") %>' />
                <br />
                NoneBusiKMS_OA:
                <asp:TextBox ID="NoneBusiKMS_OATextBox" runat="server" Text='<%# Bind("NoneBusiKMS_OA") %>' />
                <br />
                NoneBusiKMS_EA:
                <asp:TextBox ID="NoneBusiKMS_EATextBox" runat="server" Text='<%# Bind("NoneBusiKMS_EA") %>' />
                <br />
                NoneBusiKMS_OB:
                <asp:TextBox ID="NoneBusiKMS_OBTextBox" runat="server" Text='<%# Bind("NoneBusiKMS_OB") %>' />
                <br />
                NoneBusiKMS_EB:
                <asp:TextBox ID="NoneBusiKMS_EBTextBox" runat="server" Text='<%# Bind("NoneBusiKMS_EB") %>' />
                <br />
                NoneBusiKMS_OT:
                <asp:TextBox ID="NoneBusiKMS_OTTextBox" runat="server" Text='<%# Bind("NoneBusiKMS_OT") %>' />
                <br />
                NoneBusiKMS_ET:
                <asp:TextBox ID="NoneBusiKMS_ETTextBox" runat="server" Text='<%# Bind("NoneBusiKMS_ET") %>' />
                <br />
                KMS_NoneBusi:
                <asp:TextBox ID="KMS_NoneBusiTextBox" runat="server" Text='<%# Bind("KMS_NoneBusi") %>' />
                <br />
                TotalKMS_OA:
                <asp:TextBox ID="TotalKMS_OATextBox" runat="server" Text='<%# Bind("TotalKMS_OA") %>' />
                <br />
                TotalKMS_EA:
                <asp:TextBox ID="TotalKMS_EATextBox" runat="server" Text='<%# Bind("TotalKMS_EA") %>' />
                <br />
                TotalKMS_OB:
                <asp:TextBox ID="TotalKMS_OBTextBox" runat="server" Text='<%# Bind("TotalKMS_OB") %>' />
                <br />
                TotalKMS_EB:
                <asp:TextBox ID="TotalKMS_EBTextBox" runat="server" Text='<%# Bind("TotalKMS_EB") %>' />
                <br />
                TotalKMS_OT:
                <asp:TextBox ID="TotalKMS_OTTextBox" runat="server" Text='<%# Bind("TotalKMS_OT") %>' />
                <br />
                TotalKMS_ET:
                <asp:TextBox ID="TotalKMS_ETTextBox" runat="server" Text='<%# Bind("TotalKMS_ET") %>' />
                <br />
                KMS_Total:
                <asp:TextBox ID="KMS_TotalTextBox" runat="server" Text='<%# Bind("KMS_Total") %>' />
                <br />
                HighWayCount_OA:
                <asp:TextBox ID="HighWayCount_OATextBox" runat="server" Text='<%# Bind("HighWayCount_OA") %>' />
                <br />
                HighWayCount_EA:
                <asp:TextBox ID="HighWayCount_EATextBox" runat="server" Text='<%# Bind("HighWayCount_EA") %>' />
                <br />
                HighWayCount_OB:
                <asp:TextBox ID="HighWayCount_OBTextBox" runat="server" Text='<%# Bind("HighWayCount_OB") %>' />
                <br />
                HighWayCount_EB:
                <asp:TextBox ID="HighWayCount_EBTextBox" runat="server" Text='<%# Bind("HighWayCount_EB") %>' />
                <br />
                HighWayCount_OT:
                <asp:TextBox ID="HighWayCount_OTTextBox" runat="server" Text='<%# Bind("HighWayCount_OT") %>' />
                <br />
                HighWayCount_ET:
                <asp:TextBox ID="HighWayCount_ETTextBox" runat="server" Text='<%# Bind("HighWayCount_ET") %>' />
                <br />
                OtherCount_OA:
                <asp:TextBox ID="OtherCount_OATextBox" runat="server" Text='<%# Bind("OtherCount_OA") %>' />
                <br />
                OtherCount_EA:
                <asp:TextBox ID="OtherCount_EATextBox" runat="server" Text='<%# Bind("OtherCount_EA") %>' />
                <br />
                OtherCount_OB:
                <asp:TextBox ID="OtherCount_OBTextBox" runat="server" Text='<%# Bind("OtherCount_OB") %>' />
                <br />
                OtherCount_EB:
                <asp:TextBox ID="OtherCount_EBTextBox" runat="server" Text='<%# Bind("OtherCount_EB") %>' />
                <br />
                OtherCount_OT:
                <asp:TextBox ID="OtherCount_OTTextBox" runat="server" Text='<%# Bind("OtherCount_OT") %>' />
                <br />
                OtherCount_ET:
                <asp:TextBox ID="OtherCount_ETTextBox" runat="server" Text='<%# Bind("OtherCount_ET") %>' />
                <br />
                ModifyKMS:
                <asp:TextBox ID="ModifyKMSTextBox" runat="server" Text='<%# Bind("ModifyKMS") %>' />
                <br />
                SpecIncome:
                <asp:TextBox ID="SpecIncomeTextBox" runat="server" Text='<%# Bind("SpecIncome") %>' />
                <br />
                UsedCars:
                <asp:TextBox ID="UsedCarsTextBox" runat="server" Text='<%# Bind("UsedCars") %>' />
                <br />
                ElecCardCount:
                <asp:TextBox ID="ElecCardCountTextBox" runat="server" Text='<%# Bind("ElecCardCount") %>' />
                <br />
                EmpCount:
                <asp:TextBox ID="EmpCountTextBox" runat="server" Text='<%# Bind("EmpCount") %>' />
                <br />
                DriverOD:
                <asp:TextBox ID="DriverODTextBox" runat="server" Text='<%# Bind("DriverOD") %>' />
                <br />
                EmpCount_M:
                <asp:TextBox ID="EmpCount_MTextBox" runat="server" Text='<%# Bind("EmpCount_M") %>' />
                <br />
                EmpCount_F:
                <asp:TextBox ID="EmpCount_FTextBox" runat="server" Text='<%# Bind("EmpCount_F") %>' />
                <br />
                EmpCount_D:
                <asp:TextBox ID="EmpCount_DTextBox" runat="server" Text='<%# Bind("EmpCount_D") %>' />
                <br />
                BreakDown:
                <asp:TextBox ID="BreakDownTextBox" runat="server" Text='<%# Bind("BreakDown") %>' />
                <br />
                <asp:LinkButton ID="UpdateButton" runat="server" CausesValidation="True" CommandName="Update" Text="更新" />
                &nbsp;<asp:LinkButton ID="UpdateCancelButton" runat="server" CausesValidation="False" CommandName="Cancel" Text="取消" />
            </EditItemTemplate>
            <InsertItemTemplate>
                DepNoYM:
                <asp:TextBox ID="DepNoYMTextBox" runat="server" Text='<%# Bind("DepNoYM") %>' />
                <br />
                CalYM:
                <asp:TextBox ID="CalYMTextBox" runat="server" Text='<%# Bind("CalYM") %>' />
                <br />
                DepNo:
                <asp:TextBox ID="DepNoTextBox" runat="server" Text='<%# Bind("DepNo") %>' />
                <br />
                depName:
                <asp:TextBox ID="depNameTextBox" runat="server" Text='<%# Bind("depName") %>' />
                <br />
                ModifyYM:
                <asp:TextBox ID="ModifyYMTextBox" runat="server" Text='<%# Bind("ModifyYM") %>' />
                <br />
                ModifyMan:
                <asp:TextBox ID="ModifyManTextBox" runat="server" Text='<%# Bind("ModifyMan") %>' />
                <br />
                CarCount:
                <asp:TextBox ID="CarCountTextBox" runat="server" Text='<%# Bind("CarCount") %>' />
                <br />
                DriveRange:
                <asp:TextBox ID="DriveRangeTextBox" runat="server" Text='<%# Bind("DriveRange") %>' />
                <br />
                Remark:
                <asp:TextBox ID="RemarkTextBox" runat="server" Text='<%# Bind("Remark") %>' />
                <br />
                KMS_BUS:
                <asp:TextBox ID="KMS_BUSTextBox" runat="server" Text='<%# Bind("KMS_BUS") %>' />
                <br />
                KMS_Rent:
                <asp:TextBox ID="KMS_RentTextBox" runat="server" Text='<%# Bind("KMS_Rent") %>' />
                <br />
                KMS_Tour:
                <asp:TextBox ID="KMS_TourTextBox" runat="server" Text='<%# Bind("KMS_Tour") %>' />
                <br />
                KMS_Trans:
                <asp:TextBox ID="KMS_TransTextBox" runat="server" Text='<%# Bind("KMS_Trans") %>' />
                <br />
                KMS_Spec:
                <asp:TextBox ID="KMS_SpecTextBox" runat="server" Text='<%# Bind("KMS_Spec") %>' />
                <br />
                RuleCount_OA:
                <asp:TextBox ID="RuleCount_OATextBox" runat="server" Text='<%# Bind("RuleCount_OA") %>' />
                <br />
                RuleCount_EA:
                <asp:TextBox ID="RuleCount_EATextBox" runat="server" Text='<%# Bind("RuleCount_EA") %>' />
                <br />
                RuleCount_OB:
                <asp:TextBox ID="RuleCount_OBTextBox" runat="server" Text='<%# Bind("RuleCount_OB") %>' />
                <br />
                RuleCount_EB:
                <asp:TextBox ID="RuleCount_EBTextBox" runat="server" Text='<%# Bind("RuleCount_EB") %>' />
                <br />
                RuleCount_OT:
                <asp:TextBox ID="RuleCount_OTTextBox" runat="server" Text='<%# Bind("RuleCount_OT") %>' />
                <br />
                RuleCount_ET:
                <asp:TextBox ID="RuleCount_ETTextBox" runat="server" Text='<%# Bind("RuleCount_ET") %>' />
                <br />
                BusCount_OA:
                <asp:TextBox ID="BusCount_OATextBox" runat="server" Text='<%# Bind("BusCount_OA") %>' />
                <br />
                BusCount_EA:
                <asp:TextBox ID="BusCount_EATextBox" runat="server" Text='<%# Bind("BusCount_EA") %>' />
                <br />
                BusCount_OB:
                <asp:TextBox ID="BusCount_OBTextBox" runat="server" Text='<%# Bind("BusCount_OB") %>' />
                <br />
                BusCount_EB:
                <asp:TextBox ID="BusCount_EBTextBox" runat="server" Text='<%# Bind("BusCount_EB") %>' />
                <br />
                BusCount_OT:
                <asp:TextBox ID="BusCount_OTTextBox" runat="server" Text='<%# Bind("BusCount_OT") %>' />
                <br />
                BusCount_ET:
                <asp:TextBox ID="BusCount_ETTextBox" runat="server" Text='<%# Bind("BusCount_ET") %>' />
                <br />
                TourCount_OA:
                <asp:TextBox ID="TourCount_OATextBox" runat="server" Text='<%# Bind("TourCount_OA") %>' />
                <br />
                TourCount_EA:
                <asp:TextBox ID="TourCount_EATextBox" runat="server" Text='<%# Bind("TourCount_EA") %>' />
                <br />
                TourCount_OB:
                <asp:TextBox ID="TourCount_OBTextBox" runat="server" Text='<%# Bind("TourCount_OB") %>' />
                <br />
                TourCount_EB:
                <asp:TextBox ID="TourCount_EBTextBox" runat="server" Text='<%# Bind("TourCount_EB") %>' />
                <br />
                TourCount_OT:
                <asp:TextBox ID="TourCount_OTTextBox" runat="server" Text='<%# Bind("TourCount_OT") %>' />
                <br />
                TourCount_ET:
                <asp:TextBox ID="TourCount_ETTextBox" runat="server" Text='<%# Bind("TourCount_ET") %>' />
                <br />
                CarSum_OA:
                <asp:TextBox ID="CarSum_OATextBox" runat="server" Text='<%# Bind("CarSum_OA") %>' />
                <br />
                CarSum_EA:
                <asp:TextBox ID="CarSum_EATextBox" runat="server" Text='<%# Bind("CarSum_EA") %>' />
                <br />
                CarSum_OB:
                <asp:TextBox ID="CarSum_OBTextBox" runat="server" Text='<%# Bind("CarSum_OB") %>' />
                <br />
                CarSum_EB:
                <asp:TextBox ID="CarSum_EBTextBox" runat="server" Text='<%# Bind("CarSum_EB") %>' />
                <br />
                CarSum_OT:
                <asp:TextBox ID="CarSum_OTTextBox" runat="server" Text='<%# Bind("CarSum_OT") %>' />
                <br />
                CarSum_ET:
                <asp:TextBox ID="CarSum_ETTextBox" runat="server" Text='<%# Bind("CarSum_ET") %>' />
                <br />
                RuleKMS_OA:
                <asp:TextBox ID="RuleKMS_OATextBox" runat="server" Text='<%# Bind("RuleKMS_OA") %>' />
                <br />
                RuleKMS_EA:
                <asp:TextBox ID="RuleKMS_EATextBox" runat="server" Text='<%# Bind("RuleKMS_EA") %>' />
                <br />
                RuleKMS_OB:
                <asp:TextBox ID="RuleKMS_OBTextBox" runat="server" Text='<%# Bind("RuleKMS_OB") %>' />
                <br />
                RuleKMS_EB:
                <asp:TextBox ID="RuleKMS_EBTextBox" runat="server" Text='<%# Bind("RuleKMS_EB") %>' />
                <br />
                RuleKMS_OT:
                <asp:TextBox ID="RuleKMS_OTTextBox" runat="server" Text='<%# Bind("RuleKMS_OT") %>' />
                <br />
                RuleKMS_ET:
                <asp:TextBox ID="RuleKMS_ETTextBox" runat="server" Text='<%# Bind("RuleKMS_ET") %>' />
                <br />
                KMS_Rule:
                <asp:TextBox ID="KMS_RuleTextBox" runat="server" Text='<%# Bind("KMS_Rule") %>' />
                <br />
                BusKMS_OA:
                <asp:TextBox ID="BusKMS_OATextBox" runat="server" Text='<%# Bind("BusKMS_OA") %>' />
                <br />
                BusKMS_EA:
                <asp:TextBox ID="BusKMS_EATextBox" runat="server" Text='<%# Bind("BusKMS_EA") %>' />
                <br />
                BusKMS_OB:
                <asp:TextBox ID="BusKMS_OBTextBox" runat="server" Text='<%# Bind("BusKMS_OB") %>' />
                <br />
                BusKMS_EB:
                <asp:TextBox ID="BusKMS_EBTextBox" runat="server" Text='<%# Bind("BusKMS_EB") %>' />
                <br />
                BusKMS_OT:
                <asp:TextBox ID="BusKMS_OTTextBox" runat="server" Text='<%# Bind("BusKMS_OT") %>' />
                <br />
                BusKMS_ET:
                <asp:TextBox ID="BusKMS_ETTextBox" runat="server" Text='<%# Bind("BusKMS_ET") %>' />
                <br />
                SpecKMS_OA:
                <asp:TextBox ID="SpecKMS_OATextBox" runat="server" Text='<%# Bind("SpecKMS_OA") %>' />
                <br />
                SpecKMS_EA:
                <asp:TextBox ID="SpecKMS_EATextBox" runat="server" Text='<%# Bind("SpecKMS_EA") %>' />
                <br />
                SpecKMS_OB:
                <asp:TextBox ID="SpecKMS_OBTextBox" runat="server" Text='<%# Bind("SpecKMS_OB") %>' />
                <br />
                SpecKMS_EB:
                <asp:TextBox ID="SpecKMS_EBTextBox" runat="server" Text='<%# Bind("SpecKMS_EB") %>' />
                <br />
                SpecKMS_OT:
                <asp:TextBox ID="SpecKMS_OTTextBox" runat="server" Text='<%# Bind("SpecKMS_OT") %>' />
                <br />
                SpecKMS_ET:
                <asp:TextBox ID="SpecKMS_ETTextBox" runat="server" Text='<%# Bind("SpecKMS_ET") %>' />
                <br />
                RentKMS_OA:
                <asp:TextBox ID="RentKMS_OATextBox" runat="server" Text='<%# Bind("RentKMS_OA") %>' />
                <br />
                RentKMS_EA:
                <asp:TextBox ID="RentKMS_EATextBox" runat="server" Text='<%# Bind("RentKMS_EA") %>' />
                <br />
                RentKMS_OB:
                <asp:TextBox ID="RentKMS_OBTextBox" runat="server" Text='<%# Bind("RentKMS_OB") %>' />
                <br />
                RentKMS_EB:
                <asp:TextBox ID="RentKMS_EBTextBox" runat="server" Text='<%# Bind("RentKMS_EB") %>' />
                <br />
                RentKMS_OT:
                <asp:TextBox ID="RentKMS_OTTextBox" runat="server" Text='<%# Bind("RentKMS_OT") %>' />
                <br />
                RentKMS_ET:
                <asp:TextBox ID="RentKMS_ETTextBox" runat="server" Text='<%# Bind("RentKMS_ET") %>' />
                <br />
                TransKMS_OA:
                <asp:TextBox ID="TransKMS_OATextBox" runat="server" Text='<%# Bind("TransKMS_OA") %>' />
                <br />
                TransKMS_EA:
                <asp:TextBox ID="TransKMS_EATextBox" runat="server" Text='<%# Bind("TransKMS_EA") %>' />
                <br />
                TransKMS_OB:
                <asp:TextBox ID="TransKMS_OBTextBox" runat="server" Text='<%# Bind("TransKMS_OB") %>' />
                <br />
                TransKMS_EB:
                <asp:TextBox ID="TransKMS_EBTextBox" runat="server" Text='<%# Bind("TransKMS_EB") %>' />
                <br />
                TransKMS_OT:
                <asp:TextBox ID="TransKMS_OTTextBox" runat="server" Text='<%# Bind("TransKMS_OT") %>' />
                <br />
                TransKMS_ET:
                <asp:TextBox ID="TransKMS_ETTextBox" runat="server" Text='<%# Bind("TransKMS_ET") %>' />
                <br />
                TourKMS_OA:
                <asp:TextBox ID="TourKMS_OATextBox" runat="server" Text='<%# Bind("TourKMS_OA") %>' />
                <br />
                TourKMS_EA:
                <asp:TextBox ID="TourKMS_EATextBox" runat="server" Text='<%# Bind("TourKMS_EA") %>' />
                <br />
                TourKMS_OB:
                <asp:TextBox ID="TourKMS_OBTextBox" runat="server" Text='<%# Bind("TourKMS_OB") %>' />
                <br />
                TourKMS_EB:
                <asp:TextBox ID="TourKMS_EBTextBox" runat="server" Text='<%# Bind("TourKMS_EB") %>' />
                <br />
                TourKMS_OT:
                <asp:TextBox ID="TourKMS_OTTextBox" runat="server" Text='<%# Bind("TourKMS_OT") %>' />
                <br />
                TourKMS_ET:
                <asp:TextBox ID="TourKMS_ETTextBox" runat="server" Text='<%# Bind("TourKMS_ET") %>' />
                <br />
                NoneBusiKMS_OA:
                <asp:TextBox ID="NoneBusiKMS_OATextBox" runat="server" Text='<%# Bind("NoneBusiKMS_OA") %>' />
                <br />
                NoneBusiKMS_EA:
                <asp:TextBox ID="NoneBusiKMS_EATextBox" runat="server" Text='<%# Bind("NoneBusiKMS_EA") %>' />
                <br />
                NoneBusiKMS_OB:
                <asp:TextBox ID="NoneBusiKMS_OBTextBox" runat="server" Text='<%# Bind("NoneBusiKMS_OB") %>' />
                <br />
                NoneBusiKMS_EB:
                <asp:TextBox ID="NoneBusiKMS_EBTextBox" runat="server" Text='<%# Bind("NoneBusiKMS_EB") %>' />
                <br />
                NoneBusiKMS_OT:
                <asp:TextBox ID="NoneBusiKMS_OTTextBox" runat="server" Text='<%# Bind("NoneBusiKMS_OT") %>' />
                <br />
                NoneBusiKMS_ET:
                <asp:TextBox ID="NoneBusiKMS_ETTextBox" runat="server" Text='<%# Bind("NoneBusiKMS_ET") %>' />
                <br />
                KMS_NoneBusi:
                <asp:TextBox ID="KMS_NoneBusiTextBox" runat="server" Text='<%# Bind("KMS_NoneBusi") %>' />
                <br />
                TotalKMS_OA:
                <asp:TextBox ID="TotalKMS_OATextBox" runat="server" Text='<%# Bind("TotalKMS_OA") %>' />
                <br />
                TotalKMS_EA:
                <asp:TextBox ID="TotalKMS_EATextBox" runat="server" Text='<%# Bind("TotalKMS_EA") %>' />
                <br />
                TotalKMS_OB:
                <asp:TextBox ID="TotalKMS_OBTextBox" runat="server" Text='<%# Bind("TotalKMS_OB") %>' />
                <br />
                TotalKMS_EB:
                <asp:TextBox ID="TotalKMS_EBTextBox" runat="server" Text='<%# Bind("TotalKMS_EB") %>' />
                <br />
                TotalKMS_OT:
                <asp:TextBox ID="TotalKMS_OTTextBox" runat="server" Text='<%# Bind("TotalKMS_OT") %>' />
                <br />
                TotalKMS_ET:
                <asp:TextBox ID="TotalKMS_ETTextBox" runat="server" Text='<%# Bind("TotalKMS_ET") %>' />
                <br />
                KMS_Total:
                <asp:TextBox ID="KMS_TotalTextBox" runat="server" Text='<%# Bind("KMS_Total") %>' />
                <br />
                HighWayCount_OA:
                <asp:TextBox ID="HighWayCount_OATextBox" runat="server" Text='<%# Bind("HighWayCount_OA") %>' />
                <br />
                HighWayCount_EA:
                <asp:TextBox ID="HighWayCount_EATextBox" runat="server" Text='<%# Bind("HighWayCount_EA") %>' />
                <br />
                HighWayCount_OB:
                <asp:TextBox ID="HighWayCount_OBTextBox" runat="server" Text='<%# Bind("HighWayCount_OB") %>' />
                <br />
                HighWayCount_EB:
                <asp:TextBox ID="HighWayCount_EBTextBox" runat="server" Text='<%# Bind("HighWayCount_EB") %>' />
                <br />
                HighWayCount_OT:
                <asp:TextBox ID="HighWayCount_OTTextBox" runat="server" Text='<%# Bind("HighWayCount_OT") %>' />
                <br />
                HighWayCount_ET:
                <asp:TextBox ID="HighWayCount_ETTextBox" runat="server" Text='<%# Bind("HighWayCount_ET") %>' />
                <br />
                OtherCount_OA:
                <asp:TextBox ID="OtherCount_OATextBox" runat="server" Text='<%# Bind("OtherCount_OA") %>' />
                <br />
                OtherCount_EA:
                <asp:TextBox ID="OtherCount_EATextBox" runat="server" Text='<%# Bind("OtherCount_EA") %>' />
                <br />
                OtherCount_OB:
                <asp:TextBox ID="OtherCount_OBTextBox" runat="server" Text='<%# Bind("OtherCount_OB") %>' />
                <br />
                OtherCount_EB:
                <asp:TextBox ID="OtherCount_EBTextBox" runat="server" Text='<%# Bind("OtherCount_EB") %>' />
                <br />
                OtherCount_OT:
                <asp:TextBox ID="OtherCount_OTTextBox" runat="server" Text='<%# Bind("OtherCount_OT") %>' />
                <br />
                OtherCount_ET:
                <asp:TextBox ID="OtherCount_ETTextBox" runat="server" Text='<%# Bind("OtherCount_ET") %>' />
                <br />
                ModifyKMS:
                <asp:TextBox ID="ModifyKMSTextBox" runat="server" Text='<%# Bind("ModifyKMS") %>' />
                <br />
                SpecIncome:
                <asp:TextBox ID="SpecIncomeTextBox" runat="server" Text='<%# Bind("SpecIncome") %>' />
                <br />
                UsedCars:
                <asp:TextBox ID="UsedCarsTextBox" runat="server" Text='<%# Bind("UsedCars") %>' />
                <br />
                ElecCardCount:
                <asp:TextBox ID="ElecCardCountTextBox" runat="server" Text='<%# Bind("ElecCardCount") %>' />
                <br />
                EmpCount:
                <asp:TextBox ID="EmpCountTextBox" runat="server" Text='<%# Bind("EmpCount") %>' />
                <br />
                DriverOD:
                <asp:TextBox ID="DriverODTextBox" runat="server" Text='<%# Bind("DriverOD") %>' />
                <br />
                EmpCount_M:
                <asp:TextBox ID="EmpCount_MTextBox" runat="server" Text='<%# Bind("EmpCount_M") %>' />
                <br />
                EmpCount_F:
                <asp:TextBox ID="EmpCount_FTextBox" runat="server" Text='<%# Bind("EmpCount_F") %>' />
                <br />
                EmpCount_D:
                <asp:TextBox ID="EmpCount_DTextBox" runat="server" Text='<%# Bind("EmpCount_D") %>' />
                <br />
                BreakDown:
                <asp:TextBox ID="BreakDownTextBox" runat="server" Text='<%# Bind("BreakDown") %>' />
                <br />
                <asp:LinkButton ID="InsertButton" runat="server" CausesValidation="True" CommandName="Insert" Text="插入" />
                &nbsp;<asp:LinkButton ID="InsertCancelButton" runat="server" CausesValidation="False" CommandName="Cancel" Text="取消" />
            </InsertItemTemplate>
            <ItemTemplate>
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
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbTitle_4" runat="server" CssClass="titleText-S-Blue" Text="站車輛數總計" Width="100%" />
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
                DriveRange:
                <asp:Label ID="DriveRangeLabel" runat="server" Text='<%# Eval("DriveRange") %>' />
                <br />
                Remark:
                <asp:Label ID="RemarkLabel" runat="server" Text='<%# Eval("Remark") %>' />
                <br />
                KMS_BUS:
                <asp:Label ID="KMS_BUSLabel" runat="server" Text='<%# Eval("KMS_BUS") %>' />
                <br />
                KMS_Rent:
                <asp:Label ID="KMS_RentLabel" runat="server" Text='<%# Eval("KMS_Rent") %>' />
                <br />
                KMS_Tour:
                <asp:Label ID="KMS_TourLabel" runat="server" Text='<%# Eval("KMS_Tour") %>' />
                <br />
                KMS_Trans:
                <asp:Label ID="KMS_TransLabel" runat="server" Text='<%# Eval("KMS_Trans") %>' />
                <br />
                KMS_Spec:
                <asp:Label ID="KMS_SpecLabel" runat="server" Text='<%# Eval("KMS_Spec") %>' />
                <br />
                RuleCount_OB:
                <asp:Label ID="RuleCount_OBLabel" runat="server" Text='<%# Eval("RuleCount_OB") %>' />
                <br />
                RuleCount_EB:
                <asp:Label ID="RuleCount_EBLabel" runat="server" Text='<%# Eval("RuleCount_EB") %>' />
                <br />
                RuleCount_OT:
                <asp:Label ID="RuleCount_OTLabel" runat="server" Text='<%# Eval("RuleCount_OT") %>' />
                <br />
                RuleCount_ET:
                <asp:Label ID="RuleCount_ETLabel" runat="server" Text='<%# Eval("RuleCount_ET") %>' />
                <br />
                BusCount_OB:
                <asp:Label ID="BusCount_OBLabel" runat="server" Text='<%# Eval("BusCount_OB") %>' />
                <br />
                BusCount_EB:
                <asp:Label ID="BusCount_EBLabel" runat="server" Text='<%# Eval("BusCount_EB") %>' />
                <br />
                BusCount_OT:
                <asp:Label ID="BusCount_OTLabel" runat="server" Text='<%# Eval("BusCount_OT") %>' />
                <br />
                BusCount_ET:
                <asp:Label ID="BusCount_ETLabel" runat="server" Text='<%# Eval("BusCount_ET") %>' />
                <br />
                TourCount_OB:
                <asp:Label ID="TourCount_OBLabel" runat="server" Text='<%# Eval("TourCount_OB") %>' />
                <br />
                TourCount_EB:
                <asp:Label ID="TourCount_EBLabel" runat="server" Text='<%# Eval("TourCount_EB") %>' />
                <br />
                TourCount_OT:
                <asp:Label ID="TourCount_OTLabel" runat="server" Text='<%# Eval("TourCount_OT") %>' />
                <br />
                TourCount_ET:
                <asp:Label ID="TourCount_ETLabel" runat="server" Text='<%# Eval("TourCount_ET") %>' />
                <br />
                CarSum_OB:
                <asp:Label ID="CarSum_OBLabel" runat="server" Text='<%# Eval("CarSum_OB") %>' />
                <br />
                CarSum_EB:
                <asp:Label ID="CarSum_EBLabel" runat="server" Text='<%# Eval("CarSum_EB") %>' />
                <br />
                CarSum_OT:
                <asp:Label ID="CarSum_OTLabel" runat="server" Text='<%# Eval("CarSum_OT") %>' />
                <br />
                CarSum_ET:
                <asp:Label ID="CarSum_ETLabel" runat="server" Text='<%# Eval("CarSum_ET") %>' />
                <br />
                RuleKMS_OA:
                <asp:Label ID="RuleKMS_OALabel" runat="server" Text='<%# Eval("RuleKMS_OA") %>' />
                <br />
                RuleKMS_EA:
                <asp:Label ID="RuleKMS_EALabel" runat="server" Text='<%# Eval("RuleKMS_EA") %>' />
                <br />
                RuleKMS_OB:
                <asp:Label ID="RuleKMS_OBLabel" runat="server" Text='<%# Eval("RuleKMS_OB") %>' />
                <br />
                RuleKMS_EB:
                <asp:Label ID="RuleKMS_EBLabel" runat="server" Text='<%# Eval("RuleKMS_EB") %>' />
                <br />
                RuleKMS_OT:
                <asp:Label ID="RuleKMS_OTLabel" runat="server" Text='<%# Eval("RuleKMS_OT") %>' />
                <br />
                RuleKMS_ET:
                <asp:Label ID="RuleKMS_ETLabel" runat="server" Text='<%# Eval("RuleKMS_ET") %>' />
                <br />
                KMS_Rule:
                <asp:Label ID="KMS_RuleLabel" runat="server" Text='<%# Eval("KMS_Rule") %>' />
                <br />
                BusKMS_OA:
                <asp:Label ID="BusKMS_OALabel" runat="server" Text='<%# Eval("BusKMS_OA") %>' />
                <br />
                BusKMS_EA:
                <asp:Label ID="BusKMS_EALabel" runat="server" Text='<%# Eval("BusKMS_EA") %>' />
                <br />
                BusKMS_OB:
                <asp:Label ID="BusKMS_OBLabel" runat="server" Text='<%# Eval("BusKMS_OB") %>' />
                <br />
                BusKMS_EB:
                <asp:Label ID="BusKMS_EBLabel" runat="server" Text='<%# Eval("BusKMS_EB") %>' />
                <br />
                BusKMS_OT:
                <asp:Label ID="BusKMS_OTLabel" runat="server" Text='<%# Eval("BusKMS_OT") %>' />
                <br />
                BusKMS_ET:
                <asp:Label ID="BusKMS_ETLabel" runat="server" Text='<%# Eval("BusKMS_ET") %>' />
                <br />
                SpecKMS_OA:
                <asp:Label ID="SpecKMS_OALabel" runat="server" Text='<%# Eval("SpecKMS_OA") %>' />
                <br />
                SpecKMS_EA:
                <asp:Label ID="SpecKMS_EALabel" runat="server" Text='<%# Eval("SpecKMS_EA") %>' />
                <br />
                SpecKMS_OB:
                <asp:Label ID="SpecKMS_OBLabel" runat="server" Text='<%# Eval("SpecKMS_OB") %>' />
                <br />
                SpecKMS_EB:
                <asp:Label ID="SpecKMS_EBLabel" runat="server" Text='<%# Eval("SpecKMS_EB") %>' />
                <br />
                SpecKMS_OT:
                <asp:Label ID="SpecKMS_OTLabel" runat="server" Text='<%# Eval("SpecKMS_OT") %>' />
                <br />
                SpecKMS_ET:
                <asp:Label ID="SpecKMS_ETLabel" runat="server" Text='<%# Eval("SpecKMS_ET") %>' />
                <br />
                RentKMS_OA:
                <asp:Label ID="RentKMS_OALabel" runat="server" Text='<%# Eval("RentKMS_OA") %>' />
                <br />
                RentKMS_EA:
                <asp:Label ID="RentKMS_EALabel" runat="server" Text='<%# Eval("RentKMS_EA") %>' />
                <br />
                RentKMS_OB:
                <asp:Label ID="RentKMS_OBLabel" runat="server" Text='<%# Eval("RentKMS_OB") %>' />
                <br />
                RentKMS_EB:
                <asp:Label ID="RentKMS_EBLabel" runat="server" Text='<%# Eval("RentKMS_EB") %>' />
                <br />
                RentKMS_OT:
                <asp:Label ID="RentKMS_OTLabel" runat="server" Text='<%# Eval("RentKMS_OT") %>' />
                <br />
                RentKMS_ET:
                <asp:Label ID="RentKMS_ETLabel" runat="server" Text='<%# Eval("RentKMS_ET") %>' />
                <br />
                TransKMS_OA:
                <asp:Label ID="TransKMS_OALabel" runat="server" Text='<%# Eval("TransKMS_OA") %>' />
                <br />
                TransKMS_EA:
                <asp:Label ID="TransKMS_EALabel" runat="server" Text='<%# Eval("TransKMS_EA") %>' />
                <br />
                TransKMS_OB:
                <asp:Label ID="TransKMS_OBLabel" runat="server" Text='<%# Eval("TransKMS_OB") %>' />
                <br />
                TransKMS_EB:
                <asp:Label ID="TransKMS_EBLabel" runat="server" Text='<%# Eval("TransKMS_EB") %>' />
                <br />
                TransKMS_OT:
                <asp:Label ID="TransKMS_OTLabel" runat="server" Text='<%# Eval("TransKMS_OT") %>' />
                <br />
                TransKMS_ET:
                <asp:Label ID="TransKMS_ETLabel" runat="server" Text='<%# Eval("TransKMS_ET") %>' />
                <br />
                TourKMS_OA:
                <asp:Label ID="TourKMS_OALabel" runat="server" Text='<%# Eval("TourKMS_OA") %>' />
                <br />
                TourKMS_EA:
                <asp:Label ID="TourKMS_EALabel" runat="server" Text='<%# Eval("TourKMS_EA") %>' />
                <br />
                TourKMS_OB:
                <asp:Label ID="TourKMS_OBLabel" runat="server" Text='<%# Eval("TourKMS_OB") %>' />
                <br />
                TourKMS_EB:
                <asp:Label ID="TourKMS_EBLabel" runat="server" Text='<%# Eval("TourKMS_EB") %>' />
                <br />
                TourKMS_OT:
                <asp:Label ID="TourKMS_OTLabel" runat="server" Text='<%# Eval("TourKMS_OT") %>' />
                <br />
                TourKMS_ET:
                <asp:Label ID="TourKMS_ETLabel" runat="server" Text='<%# Eval("TourKMS_ET") %>' />
                <br />
                NoneBusiKMS_OA:
                <asp:Label ID="NoneBusiKMS_OALabel" runat="server" Text='<%# Eval("NoneBusiKMS_OA") %>' />
                <br />
                NoneBusiKMS_EA:
                <asp:Label ID="NoneBusiKMS_EALabel" runat="server" Text='<%# Eval("NoneBusiKMS_EA") %>' />
                <br />
                NoneBusiKMS_OB:
                <asp:Label ID="NoneBusiKMS_OBLabel" runat="server" Text='<%# Eval("NoneBusiKMS_OB") %>' />
                <br />
                NoneBusiKMS_EB:
                <asp:Label ID="NoneBusiKMS_EBLabel" runat="server" Text='<%# Eval("NoneBusiKMS_EB") %>' />
                <br />
                NoneBusiKMS_OT:
                <asp:Label ID="NoneBusiKMS_OTLabel" runat="server" Text='<%# Eval("NoneBusiKMS_OT") %>' />
                <br />
                NoneBusiKMS_ET:
                <asp:Label ID="NoneBusiKMS_ETLabel" runat="server" Text='<%# Eval("NoneBusiKMS_ET") %>' />
                <br />
                KMS_NoneBusi:
                <asp:Label ID="KMS_NoneBusiLabel" runat="server" Text='<%# Eval("KMS_NoneBusi") %>' />
                <br />
                TotalKMS_OA:
                <asp:Label ID="TotalKMS_OALabel" runat="server" Text='<%# Eval("TotalKMS_OA") %>' />
                <br />
                TotalKMS_EA:
                <asp:Label ID="TotalKMS_EALabel" runat="server" Text='<%# Eval("TotalKMS_EA") %>' />
                <br />
                TotalKMS_OB:
                <asp:Label ID="TotalKMS_OBLabel" runat="server" Text='<%# Eval("TotalKMS_OB") %>' />
                <br />
                TotalKMS_EB:
                <asp:Label ID="TotalKMS_EBLabel" runat="server" Text='<%# Eval("TotalKMS_EB") %>' />
                <br />
                TotalKMS_OT:
                <asp:Label ID="TotalKMS_OTLabel" runat="server" Text='<%# Eval("TotalKMS_OT") %>' />
                <br />
                TotalKMS_ET:
                <asp:Label ID="TotalKMS_ETLabel" runat="server" Text='<%# Eval("TotalKMS_ET") %>' />
                <br />
                KMS_Total:
                <asp:Label ID="KMS_TotalLabel" runat="server" Text='<%# Eval("KMS_Total") %>' />
                <br />
                HighWayCount_OB:
                <asp:Label ID="HighWayCount_OBLabel" runat="server" Text='<%# Eval("HighWayCount_OB") %>' />
                <br />
                HighWayCount_EB:
                <asp:Label ID="HighWayCount_EBLabel" runat="server" Text='<%# Eval("HighWayCount_EB") %>' />
                <br />
                HighWayCount_OT:
                <asp:Label ID="HighWayCount_OTLabel" runat="server" Text='<%# Eval("HighWayCount_OT") %>' />
                <br />
                HighWayCount_ET:
                <asp:Label ID="HighWayCount_ETLabel" runat="server" Text='<%# Eval("HighWayCount_ET") %>' />
                <br />
                OtherCount_OB:
                <asp:Label ID="OtherCount_OBLabel" runat="server" Text='<%# Eval("OtherCount_OB") %>' />
                <br />
                OtherCount_EB:
                <asp:Label ID="OtherCount_EBLabel" runat="server" Text='<%# Eval("OtherCount_EB") %>' />
                <br />
                OtherCount_OT:
                <asp:Label ID="OtherCount_OTLabel" runat="server" Text='<%# Eval("OtherCount_OT") %>' />
                <br />
                OtherCount_ET:
                <asp:Label ID="OtherCount_ETLabel" runat="server" Text='<%# Eval("OtherCount_ET") %>' />
                <br />
                ModifyKMS:
                <asp:Label ID="ModifyKMSLabel" runat="server" Text='<%# Eval("ModifyKMS") %>' />
                <br />
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdCarCount" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT DepNoYM, CalYM, c.DepNo, d.[Name] DepName, ModifyYM, ModifyMan, e.[Name] EmpName, CarCount, DriveRange, c.Remark, 
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
 WHERE isnull(DepNoYM, '') = ''"></asp:SqlDataSource>
</asp:Content>
