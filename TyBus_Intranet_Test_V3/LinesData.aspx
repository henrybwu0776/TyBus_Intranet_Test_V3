<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="LinesData.aspx.cs" Inherits="TyBus_Intranet_Test_V3.LinesData" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="LinesDataForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">路線基本資料</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbLinesNo_Search" runat="server" CssClass="text-Right-Blue" Text="ERP路線編號" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eLinesNoS_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="titleText-Black" Text="～" />
                    <asp:TextBox ID="eLinesNoE_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbTicketLinesNo_Search" runat="server" CssClass="text-Right-Blue" Text="驗票機路線編號" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eTicketLinesNoS_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="titleText-Black" Text="～" />
                    <asp:TextBox ID="eTicketLinesNoE_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbLinesNoM_Search" runat="server" CssClass="text-Right-Blue" Text="共線代號" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eLinesNoMS_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_3" runat="server" CssClass="titleText-Black" Text="～" />
                    <asp:TextBox ID="eLinesNoME_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbLCDLinesNo_Search" runat="server" CssClass="text-Right-Blue" Text="LCD路線編號" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eLCDLinesNoS_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_4" runat="server" CssClass="titleText-Black" Text="～" />
                    <asp:TextBox ID="eLCDLinesNoE_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbCallName_Search" runat="server" CssClass="text-Right-Blue" Text="路線簡稱" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:TextBox ID="eCallName_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbLicenseNo_Search" runat="server" CssClass="text-Right-Blue" Text="許可證號" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:TextBox ID="eLicenseNo_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbCarType_Search" runat="server" CssClass="text-Right-Blue" Text="車種" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:DropDownList ID="ddlCarType_Search" runat="server" CssClass="text-Left-Black" Width="95%" OnSelectedIndexChanged="ddlCarType_Search_SelectedIndexChanged"
                        AutoPostBack="True" DataSourceID="dsCarType_Search" DataTextField="ClassTxt" DataValueField="ClassNo" />
                    <asp:Label ID="eCarType_Search" runat="server" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbCustNo_Search" runat="server" CssClass="text-Right-Blue" Text="客戶編號" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:TextBox ID="eCustNo_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Blue" Text="查詢" OnClick="bbSearch_Click" Width="120px" />
                </td>
                <td class="ColWidth-9Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="120px" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridLines_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="LinesNo" DataSourceID="dsLines_List" ForeColor="#333333" GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="LinesNo" HeaderText="路線編號" ReadOnly="True" SortExpression="LinesNo" />
                <asp:BoundField DataField="LineName" HeaderText="路線說明" SortExpression="LineName" />
                <asp:BoundField DataField="CallName" HeaderText="路線簡稱" SortExpression="CallName" />
                <asp:BoundField DataField="LCDLinesNo" HeaderText="LCD路線編號" SortExpression="LCDLinesNo" />
                <asp:BoundField DataField="TicketLineNo" HeaderText="驗票機路線" SortExpression="TicketLineNo" />
                <asp:BoundField DataField="LinesNoM" HeaderText="共線代號" SortExpression="LinesNoM" />
                <asp:BoundField DataField="LicenseNo" HeaderText="許可證號" SortExpression="LicenseNo" />
                <asp:BoundField DataField="CarType" HeaderText="車種代碼" SortExpression="CarType" Visible="False" />
                <asp:BoundField DataField="CarType_C" HeaderText="車種" ReadOnly="True" SortExpression="CarType_C" />
                <asp:BoundField DataField="IsExtra" HeaderText="補貼路線" SortExpression="IsExtra" />
                <asp:BoundField DataField="Highway" HeaderText="國道" SortExpression="Highway" />
                <asp:BoundField DataField="CustNo" HeaderText="客戶編號" SortExpression="CustNo" Visible="False" />
                <asp:BoundField DataField="CustName" HeaderText="客戶" ReadOnly="True" SortExpression="CustName" />
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
        <asp:FormView ID="fvLines_Detail" runat="server" DataKeyNames="LinesNo" DataSourceID="dsLines_Detail" OnDataBound="fvLines_Detail_DataBound" Width="100%">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_Edit_Click" Text="確定" Width="120px" />
                &nbsp;
                <asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="UpdatePanel_Edit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNo_Edit" runat="server" CssClass="text-Right-Blue" Text="路線編號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eLinesNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesName_Edit" runat="server" CssClass="text-Right-Blue" Text="路線名稱" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eLineName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LineName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCallName_Edit" runat="server" CssClass="text-Right-Blue" Text="簡稱" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCallName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("callname") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="bTicketLineNo_Edit" runat="server" CssClass="text-Right-Blue" Text="驗票機路線編號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTicketLineNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TicketLineNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNoM_Edit" runat="server" CssClass="text-Right-Blue" Text="共線代號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLinesNoM_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNoM") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarType_Edit" runat="server" CssClass="text-Right-Blue" Text="車種" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlCarType_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlCarType_Edit_SelectedIndexChanged" Width="95%"
                                        DataSourceID="dsCarType_Edit" DataTextField="ClassTxt" DataValueField="ClassNo" />
                                    <asp:Label ID="eCarType_Edit" runat="server" Text='<%# Eval("cartype") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFixedCarCount_Edit" runat="server" CssClass="text-Right-Blue" Text="固定車輛數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eFixedCarCount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FixedCarCount") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseNo_Edit" runat="server" CssClass="text-Right-Blue" Text="許可證編號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eLicenseNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesGovNo_Edit" runat="server" CssClass="text-Right-Blue" Text="規定編碼" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eLinesGovNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesGovNo") %>' Width="55%" />
                                    <asp:TextBox ID="eLinesGOVExtNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesGOVExtNo") %>' wid="40%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbMainGovOffice_Edit" runat="server" CssClass="text-Right-Blue" Text="主管機關" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eMainGovOffice_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("MainGovOffice") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbActualRun_Edit" runat="server" CssClass="text-Right-Blue" Text="實際班次數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eActualRun_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ActualRun") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseRun_Edit" runat="server" CssClass="text-Right-Blue" Text="核定班次數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLicenseRun_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseRun") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseRunSun_Edit" runat="server" CssClass="text-Right-Blue" Text="假日核定班次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLicenseRunSun_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseRunSun") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseRunLV_Edit" runat="server" CssClass="text-Right-Blue" Text="寒暑假核定班次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLicenseRunLV_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseRunLV") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbActualKM_Edit" runat="server" CssClass="text-Right-Blue" Text="實際公里數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eActualKM_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ActualKM") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbMins_Edit" runat="server" CssClass="text-Right-Blue" Text="時間 (分鐘數)" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eMins_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Mins") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseKM_Edit" runat="server" CssClass="text-Right-Blue" Text="核定公里數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLicenseKM_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseKM") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAllowanceKM_Edit" runat="server" CssClass="text-Right-Blue" Text="補貼公里數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAllowanceKm_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("allowanceKm") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNormal_Edit" runat="server" CssClass="text-Right-Blue" Text="標準趟次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eNormal_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("normal") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="6">
                                    <asp:Label ID="lbNote_1_Edit" runat="server" CssClass="text-Left-Black" Text="超過 " />
                                    <asp:TextBox ID="eFeedKm_Edit" runat="server" CssClass="text-Right-Black" Text='<%# Eval("FeedKm") %>' Width="5%" />
                                    <asp:Label ID="lbNote_2_Edit" runat="server" CssClass="text-Left-Black" Text="公里，則" />
                                    <asp:TextBox ID="eFeedAMT2_Edit" runat="server" CssClass="text-Right-Black" Text='<%# Eval("FeedAmt2") %>' Width="5%" />
                                    <asp:Label ID="lbNote_3_Edit" runat="server" CssClass="text-Left-Black" Text="元，否則" />
                                    <asp:TextBox ID="eFeedAMT1_Edit" runat="server" CssClass="text-Right-Black" Text='<%# Eval("FeedAmt1") %>' Width="5%" />
                                    <asp:Label ID="lbNote_4_Edit" runat="server" CssClass="text-Left-Black" Text="元" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:CheckBox ID="cbHighWay_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbHighWay_Edit_CheckedChanged" Text="國道路線" Width="12%" />
                                    <asp:CheckBox ID="cbIsExtra_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbIsExtra_Edit_CheckedChanged" Text="補貼路線" Width="12%" />
                                    <asp:CheckBox ID="cbBSDBSPKM_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbBSDBSPKM_Edit_CheckedChanged" Text="不計站損益專車公里數" Width="24%" />
                                    <asp:CheckBox ID="cbOperation_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbOperation_Edit_CheckedChanged" Text="非營運" Width="12%" />
                                    <asp:CheckBox ID="cbIsFeed_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbIsFeed_Edit_CheckedChanged" Text="趟次加給" Width="12%" />
                                    <asp:CheckBox ID="cbIsCycleLine_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbIsCycleLine_Edit_CheckedChanged" Text="循環路線" Width="12%" />
                                    <asp:Label ID="eHighWay_Edit" runat="server" Text='<%# Eval("highWay") %>' Visible="false" />
                                    <asp:Label ID="eIsExtra_Edit" runat="server" Text='<%# Eval("IsExtra") %>' Visible="false" />
                                    <asp:Label ID="eBSDBSPKM_Edit" runat="server" Text='<%# Eval("BSDBSPKM") %>' Visible="false" />
                                    <asp:Label ID="eOperation_Edit" runat="server" Text='<%# Eval("operation") %>' Visible="false" />
                                    <asp:Label ID="eIsFeed_Edit" runat="server" Text='<%# Eval("IsFeed") %>' Visible="false" />
                                    <asp:Label ID="eIsCycleLine_Edit" runat="server" Text='<%# Eval("IsCycleLine") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSplit_5" runat="server" CssClass="text-Right-Blue" Text="18項成本用" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbApprovedTimes_Edit" runat="server" CssClass="text-Right-Blue" Text="核定趟次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eApprovedTimes_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ApprovedTimes") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbToll_Edit" runat="server" CssClass="text-Right-Blue" Text="通行費" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eToll_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Toll") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbApprovedDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="核定單位" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eApprovedDepNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ApprovedDepNo") %>' Width="95%" />
                                </td>    
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFixFile_Edit" runat="server" CssClass="text-Right-Blue" Text="路線附檔 (最大 4MB)" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="4">
                                    <asp:Label ID="eFixFile_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FIXFILE") %>' Width="97%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbMainDep_Edit" runat="server" CssClass="text-Right-Blue" Text="主責單位" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Textbox ID="eMainDep_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" 
                                        OnTextChanged="eMainDep_Edit_TextChanged" Text='<%# Eval("MainDep") %>' Width="30%" />
                                    <asp:Label ID="eMainDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("MainDepName") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCustNo_Edit" runat="server" CssClass="text-Right-Blue" Text="合約車客戶" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCustNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eCustNo_Edit_TextChanged" Text='<%# Eval("custno") %>' Width="35%" />
                                    <asp:Label ID="eCustName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CustName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="基價" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAmount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("amount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbUnTAxAmt_Edit" runat="server" CssClass="text-Right-Blue" Text="未稅金額" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eUnTaxAmt_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UnTaxAmt") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDateE_Edit" runat="server" CssClass="text-Right-Blue" Text="路線到期日" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDateE_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DATEE","{0:yyyy/MM/dd}") %>' Width="95%" />
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
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_INS_Click" Text="確定" Width="120px" />
                &nbsp;
                <asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="UpdatePanel_INS" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNo_INS" runat="server" CssClass="text-Right-Blue" Text="路線編號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLinesNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnTextChanged="eLinesNo_INS_TextChanged" Text='<%# Eval("LinesNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesName_INS" runat="server" CssClass="text-Right-Blue" Text="路線名稱" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eLineName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LineName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCallName_INS" runat="server" CssClass="text-Right-Blue" Text="簡稱" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCallName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("callname") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="bTicketLineNo_INS" runat="server" CssClass="text-Right-Blue" Text="驗票機路線編號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eTicketLineNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TicketLineNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNoM_INS" runat="server" CssClass="text-Right-Blue" Text="共線代號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLinesNoM_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNoM") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarType_INS" runat="server" CssClass="text-Right-Blue" Text="車種" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlCarType_INS" runat="server" CssClass="text-Left-Black" Width="95%" AutoPostBack="True" OnSelectedIndexChanged="ddlCarType_INS_SelectedIndexChanged"
                                        DataSourceID="dsCarType_INS" DataTextField="ClassTxt" DataValueField="ClassNo" />
                                    <asp:Label ID="eCarType_INS" runat="server" Text='<%# Eval("cartype") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFixedCarCount_INS" runat="server" CssClass="text-Right-Blue" Text="固定車輛數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eFixedCarCount_INS" runat="server" CssClass="text-Left-Black" Text='<%#Eval("FixedCarCount") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseNo_INS" runat="server" CssClass="text-Right-Blue" Text="許可證編號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eLicenseNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eLicenseNo_INS_TextChanged" Text='<%# Eval("LicenseNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesGovNo_INS" runat="server" CssClass="text-Right-Blue" Text="規定編碼" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eLinesGovNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesGovNo") %>' Width="55%" />
                                    <asp:TextBox ID="eLinesGIVExtNo_INS" runat="server" CssClass="text-Left-Black" Text='<%#Eval("LinesGOVExtNo") %>' wid="40%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbMainGovOffice_INS" runat="server" CssClass="text-Right-Blue" Text="主管機關" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eMainGovOffice_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("MainGovOffice") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbActualRun_INS" runat="server" CssClass="text-Right-Blue" Text="實際班次數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eActualRun_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ActualRun") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseRun_INS" runat="server" CssClass="text-Right-Blue" Text="核定班次數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLicenseRun_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseRun") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseRunSun_INS" runat="server" CssClass="text-Right-Blue" Text="假日核定班次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLicenseRunSun_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseRunSun") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseRunLV_INS" runat="server" CssClass="text-Right-Blue" Text="寒暑假核定班次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLicenseRunLV_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseRunLV") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbActualKM_INS" runat="server" CssClass="text-Right-Blue" Text="實際公里數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eActualKM_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ActualKM") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbMins_INS" runat="server" CssClass="text-Right-Blue" Text="時間 (分鐘數)" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eMins_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Mins") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseKM_INS" runat="server" CssClass="text-Right-Blue" Text="核定公里數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLicenseKM_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseKM") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAllowanceKM_INS" runat="server" CssClass="text-Right-Blue" Text="補貼公里數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAllowanceKm_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("allowanceKm") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNormal_INS" runat="server" CssClass="text-Right-Blue" Text="標準趟次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eNormal_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("normal") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="6">
                                    <asp:Label ID="lbNote_1_INS" runat="server" CssClass="text-Left-Black" Text="超過 " />
                                    <asp:TextBox ID="eFeedKM_INS" runat="server" CssClass="text-Right-Black" Text='<%# Eval("FeedKm") %>' Width="5%" />
                                    <asp:Label ID="lbNote_2_INS" runat="server" CssClass="text-Left-Black" Text="公里，則" />
                                    <asp:TextBox ID="eFeedAMT2_INS" runat="server" CssClass="text-Right-Black" Text='<%# Eval("FeedAmt2") %>' Width="5%" />
                                    <asp:Label ID="lbNote_3_INS" runat="server" CssClass="text-Left-Black" Text="元，否則" />
                                    <asp:TextBox ID="eFeedAMT1_INS" runat="server" CssClass="text-Right-Black" Text='<%# Eval("FeedAmt1") %>' Width="5%" />
                                    <asp:Label ID="lbNote_4_INS" runat="server" CssClass="text-Left-Black" Text="元" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:CheckBox ID="cbHighWay_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnCheckedChanged="cbHighWay_INS_CheckedChanged" Text="國道路線" Width="12%" />
                                    <asp:CheckBox ID="cbIsExtra_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnCheckedChanged="cbIsExtra_INS_CheckedChanged" Text="補貼路線" Width="12%" />
                                    <asp:CheckBox ID="cbBSDBSPKM_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnCheckedChanged="cbBSDBSPKM_INS_CheckedChanged" Text="不計站損益專車公里數" Width="24%" />
                                    <asp:CheckBox ID="cbOperation_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnCheckedChanged="cbOpration_INS_CheckedChanged" Text="非營運" Width="12%" />
                                    <asp:CheckBox ID="cbIsFeed_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true"
                                        OnCheckedChanged="cbIsFeed_INS_CheckedChanged" Text="趟次加給" Width="12%" />
                                     <asp:CheckBox ID="cbIsCycleLine_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" 
                                       OnCheckedChanged="cbIsCycleLine_INS_CheckedChanged" Text="循環路線" Width="12%" />
                                    <asp:Label ID="eHighWay_INS" runat="server" Text='<%# Eval("highWay") %>' Visible="false" />
                                    <asp:Label ID="eIsExtra_INS" runat="server" Text='<%# Eval("IsExtra") %>' Visible="false" />
                                    <asp:Label ID="eBSDBSPKM_INS" runat="server" Text='<%# Eval("BSDBSPKM") %>' Visible="false" />                                     
                                    <asp:Label ID="eOperation_INS" runat="server" Text='<%# Eval("operation") %>' Visible="false" />
                                    <asp:Label ID="eIsFeed_INS" runat="server" Text='<%# Eval("IsFeed") %>' Visible="false" />
                                    <asp:Label ID="eIsCycleLine_INS" runat="server" Text='<%# Eval("IsCycleLine") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSplit_7" runat="server" CssClass="text-Right-Blue" Text="18項成本用" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbApprovedTimes_INS" runat="server" CssClass="text-Right-Blue" Text="核定趟次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eApprovedTimes_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ApprovedTimes") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbToll_INS" runat="server" CssClass="text-Right-Blue" Text="通行費" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eToll_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Toll") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbApprovedDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="核定單位" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eApprovedDepNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ApprovedDepNo") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFixFile_INS" runat="server" CssClass="text-Right-Blue" Text="路線附檔 (最大 4MB)" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="4">
                                    <asp:FileUpload ID="fuFixFile_INS" runat="server" CssClass="text-Left-Blue" Width="97%" />
                                    <asp:Label ID="eFixFile_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FIXFILE") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbMainDep_INS" runat="server" CssClass="text-Right-Blue" Text="主責單位" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eMainDep_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" 
                                        OnTextChanged="eMainDep_INS_TextChanged" Text='<%# Eval("MainDep") %>' Width="30%" />
                                    <asp:Label ID="eMainDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("MainDepName") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCustNo_INS" runat="server" CssClass="text-Right-Blue" Text="合約車客戶" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCustNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("custno") %>' Width="35%" />
                                    <asp:Label ID="eCustName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CustName") %>' Width="60%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAmount_INS" runat="server" CssClass="text-Right-Blue" Text="基價" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAmount_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("amount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbUnTAxAmt_INS" runat="server" CssClass="text-Right-Blue" Text="未稅金額" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eUnTaxAmt_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UnTaxAmt") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDateE_INS" runat="server" CssClass="text-Right-Blue" Text="路線到期日" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDateE_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DATEE","{0:yyyy/MM/dd}") %>' Width="95%" />
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
                <asp:Button ID="bbNew_Empty" runat="server" CssClass="button-Blue" CausesValidation="false" CommandName="New" Text="新增" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_List" runat="server" CssClass="button-Black" CausesValidation="false" CommandName="New" Text="新增" Width="120px" />
                &nbsp;
                <asp:Button ID="bbEdit_List" runat="server" CssClass="button-Blue" CausesValidation="false" CommandName="Edit" Text="修改" Width="120px" />
                &nbsp;
                <asp:Button ID="bbDelete_List" runat="server" CssClass="button-Red" OnClick="bbDelete_List_Click" Text="刪除" Width="120px" />
                <asp:UpdatePanel ID="UpdatePanel_List" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNo_List" runat="server" CssClass="text-Right-Blue" Text="路線編號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eLinesNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesName_List" runat="server" CssClass="text-Right-Blue" Text="路線名稱" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:Label ID="eLineName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LineName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCallName_List" runat="server" CssClass="text-Right-Blue" Text="簡稱" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCallName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("callname") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="bTicketLineNo_List" runat="server" CssClass="text-Right-Blue" Text="驗票機路線編號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eTicketLineNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TicketLineNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNoM_List" runat="server" CssClass="text-Right-Blue" Text="共線代號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eLinesNoM_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNoM") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarType_List" runat="server" CssClass="text-Right-Blue" Text="車種" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCarType_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarType_C") %>' Width="95%" />
                                    <asp:Label ID="eCarType_List" runat="server" Text='<%# Eval("cartype") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFixedCarCount_List" runat="server" CssClass="text-Right-Blue" Text="固定車輛數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eFixedCarCount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FixedCarCount") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseNo_List" runat="server" CssClass="text-Right-Blue" Text="許可證編號" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eLicenseNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseNo") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesGovNo_List" runat="server" CssClass="text-Right-Blue" Text="規定編碼" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eLinesGovNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesGovNo") %>' Width="55%" />
                                    <asp:Label ID="eLinesGOVExtNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesGOVExtNo") %>' wid="40%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbMainGovOffice_List" runat="server" CssClass="text-Right-Blue" Text="主管機關" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eMainGovOffice_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("MainGovOffice") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbActualRun_List" runat="server" CssClass="text-Right-Blue" Text="實際班次數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eActualRun_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ActualRun") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseRun_List" runat="server" CssClass="text-Right-Blue" Text="核定班次數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eLicenseRun_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseRun") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseRunSun_List" runat="server" CssClass="text-Right-Blue" Text="假日核定班次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eLicenseRunSun_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseRunSun") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseRunLV_List" runat="server" CssClass="text-Right-Blue" Text="寒暑假核定班次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eLicenseRunLV_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseRunLV") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbActualKM_List" runat="server" CssClass="text-Right-Blue" Text="實際公里數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eActualKM_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ActualKM") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbMins_List" runat="server" CssClass="text-Right-Blue" Text="時間 (分鐘數)" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eMins_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Mins") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLicenseKM_List" runat="server" CssClass="text-Right-Blue" Text="核定公里數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eLicenseKM_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LicenseKM") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAllowanceKM_List" runat="server" CssClass="text-Right-Blue" Text="補貼公里數" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eAllowanceKm_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("allowanceKm") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbNormal_List" runat="server" CssClass="text-Right-Blue" Text="標準趟次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eNormal_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("normal") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="6">
                                    <asp:Label ID="lbNote_1_List" runat="server" CssClass="text-Left-Black" Text="超過 " />
                                    <asp:Label ID="eFeedKm_List" runat="server" CssClass="text-Right-Black" Text='<%# Eval("FeedKm") %>' Width="5%" />
                                    <asp:Label ID="lbNote_2_List" runat="server" CssClass="text-Left-Black" Text="公里，則" />
                                    <asp:Label ID="eFeedAMT2_List" runat="server" CssClass="text-Right-Black" Text='<%# Eval("FeedAmt2") %>' Width="5%" />
                                    <asp:Label ID="lbNote_3_List" runat="server" CssClass="text-Left-Black" Text="元，否則" />
                                    <asp:Label ID="eFeedAMT1_List" runat="server" CssClass="text-Right-Black" Text='<%# Eval("FeedAmt1") %>' Width="5%" />
                                    <asp:Label ID="lbNote_4_List" runat="server" CssClass="text-Left-Black" Text="元" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                    <asp:CheckBox ID="cbHighWay_List" runat="server" CssClass="text-Left-Black" Text="國道路線" Enabled="false" Width="12%" />
                                    <asp:CheckBox ID="cbIsExtra_List" runat="server" CssClass="text-Left-Black" Text="補貼路線" Enabled="false" Width="12%" />
                                    <asp:Label ID="eHighWay_List" runat="server" Text='<%# Eval("highWay") %>' Visible="false" />
                                    <asp:Label ID="eIsExtra_List" runat="server" Text='<%# Eval("IsExtra") %>' Visible="false" />
                                    <asp:CheckBox ID="cbBSDBSPKM_List" runat="server" CssClass="text-Left-Black" Text="不計站損益專車公里數" Enabled="false" Width="24%" />
                                    <asp:Label ID="eBSDBSPKM_List" runat="server" Text='<%# Eval("BSDBSPKM") %>' Visible="false" />
                                    <asp:CheckBox ID="cbOperation_List" runat="server" CssClass="text-Left-Black" Text="非營運" Enabled="false" Width="12%" />
                                    <asp:CheckBox ID="cbIsFeed_List" runat="server" CssClass="text-Left-Black" Text="趟次加給" Enabled="false" Width="12%" />
                                    <asp:Label ID="eOperation_List" runat="server" Text='<%# Eval("operation") %>' Visible="false" />
                                    <asp:Label ID="eIsFeed_List" runat="server" Text='<%# Eval("IsFeed") %>' Visible="false" />
                                    <asp:CheckBox ID="cbIsCycleLine_List" runat="server" CssClass="text-Left-Black" Text="循環路線" Enabled="false" Width="12%" />
                                    <asp:Label ID="eIsCycleLine_List" runat="server" Text='<%# Eval("IsCycleLine") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSplit_9" runat="server" CssClass="text-Right-Blue" Text="18項成本用" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbApprovedTimes_List" runat="server" CssClass="text-Right-Blue" Text="核定趟次" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eApprovedTimes_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ApprovedTimes") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbToll_List" runat="server" CssClass="text-Right-Blue" Text="通行費" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eToll_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Toll") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbApprovedDepNo_List" runat="server" CssClass="text-Right-Blue" Text="核定單位" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eApprovedDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ApprovedDepNo") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFixFile_List" runat="server" CssClass="text-Right-Blue" Text="路線附檔" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="4">
                                    <asp:Label ID="eFixFile_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FIXFILE") %>' Width="97%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbMainDep_List" runat="server" CssClass="text-Right-Blue" Text="主責單位" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eMainDep_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("MainDep") %>' Visible="false" />
                                    <asp:Label ID="eMainDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("MainDepName") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCustNo_List" runat="server" CssClass="text-Right-Blue" Text="合約車客戶" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCustNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("custno") %>' Visible="false" />
                                    <asp:Label ID="eCustName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CustName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAmount_List" runat="server" CssClass="text-Right-Blue" Text="基價" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("amount") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbUnTAxAmt_List" runat="server" CssClass="text-Right-Blue" Text="未稅金額" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eUnTaxAmt_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UnTaxAmt") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDateE_List" runat="server" CssClass="text-Right-Blue" Text="路線到期日" Width="100%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eDateE_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DATEE","{0:yyyy/MM/dd}") %>' Width="95%" />
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
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="dsCarType_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        SelectCommand="SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '班次表          Runs            CARTYPE') ORDER BY ClassNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsCarType_Edit" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        SelectCommand="SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '班次表          Runs            CARTYPE') ORDER BY ClassNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsCarType_INS" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        SelectCommand="SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '班次表          Runs            CARTYPE') ORDER BY ClassNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsLines_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT a.LinesNo, a.LineName, a.callname, a.LicenseNo, b.LCDLINESNO, a.TicketLineNo, a.LinesNoM, a.cartype, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '班次表          Runs            CARTYPE') AND (CLASSNO = a.cartype)) AS CarType_C, a.IsExtra, a.highWay, a.custno, (SELECT name FROM CUSTOM WHERE (types = 'C') AND (code = a.custno)) AS CustName, a.FixedCarCount, a.MainGovOffice FROM Lines AS a LEFT OUTER JOIN LinesB AS b ON b.LinesNo = a.LinesNo WHERE (ISNULL(a.LinesNo, '') &lt;&gt; '')"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsLines_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" 
        SelectCommand="SELECT LinesNo, LineName, LicenseNo, ActualRun, ActualKM, Mins, FIXFILE, DATEE, IsFeed, callname, IsExtra, LinesNoM, operation, custno, (SELECT name FROM CUSTOM WHERE (types = 'C') AND (code = a.custno)) AS CustName, amount, UnTaxAmt, allowanceKm, cartype, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '班次表          Runs            CARTYPE') AND (CLASSNO = a.cartype)) AS CarType_C, normal, FeedKm, FeedAmt1, FeedAmt2, highWay, TicketLineNo, BSDBSPKM, ApprovedTimes, Toll, ApprovedDepNo, LinesGovNo, LinesGOVExtNo, LicenseRun, LicenseRunSun, LicenseRunLV, LicenseKM, IsCycleLine, MainDep, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.MainDep)) AS MainDepName, MainGovOffice, FixedCarCount FROM Lines AS a WHERE (LinesNo = @LinesNo)" 
        DeleteCommand="DELETE FROM Lines WHERE (LinesNo = @LinesNo)" 
        InsertCommand="INSERT INTO Lines(LinesNo, LineName, callname, LinesNoM, LicenseNo, LinesGovNo, LinesGOVExtNo, cartype, TicketLineNo, ActuakRun, highWay, ActualKM, Mins, allowanceKm, IsExtra, normal, FeedKm, FeedAmt1, FeedAmt2, BSDBSPKM, ApprovedTimes, Toll, ApprovedDepNo, operation, FIXFILE, IsFeed, custno, amount, UnTaxAmt, DATEE, FixedCarCount, MainGovOffice) VALUES (@LinesNo, @LineName, @CallName, LinesNoM, @LicenseNo, @LinesGovNo, @LinesGOVExtNo, @CarType, @TicketLineNo, @ActuakRun, highWay, @ActualKM, @Mins, @AllowanceKM, @IsExtra, @Normal, @FeedKM, @FeedAMT1, @FeedAMT2, @BSDBSPKM, ApprovedTimes, @Toll, @ApprovedDepNo, @Operation, @FixFile, @IsFeed, @CustNo, @Amount, @UnTaxAMT, @DateE, @FixedCarCount, @MainGovOffice)" 
        UpdateCommand="UPDATE Lines SET LineName = @LineName, callname = @CallName, LinesNoM = @LinesNoM, LinesGovNo = @LinesGovNo, LinesGOVExtNo = @LinesGOVExtNo, cartype = @CarType, TicketLineNo = @TicketLineNo, ActualRun = @ActualRun, highWay = @HighWay, IsEctra = @IsExtra, BSDBSPKM = @BSDBSPKM, operation = @Operation, IsFeed = @IsFeed, ActualKM = @ActualKM, Mins = @Mins, allowanceKm = @AllowanceKM, normal = @Normal, FeedKm = @FeedKM, FeedAmt2 = @FeedAMT2, FeedAmt1 = @FeedAMT1, ApprovedTimes = @ApprovedTimes, Toll = @Toll, ApprovedDepNo = @ApprovedDepNo, FIXFILE = @FixFile, custno = @CustNo, amount = @Amount, UnTaxAmt = @UnTaxAMT, DATEE = @DateE, FixedCarCount = @FixedCarCount, MainGovOffice = @MainGovOffice WHERE (LinesNo = @LinesNo)">
        <DeleteParameters>
            <asp:Parameter Name="LinesNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="LinesNo" />
            <asp:Parameter Name="LineName" />
            <asp:Parameter Name="CallName" />
            <asp:Parameter Name="LicenseNo" />
            <asp:Parameter Name="LinesGovNo" />
            <asp:Parameter Name="LinesGOVExtNo" />
            <asp:Parameter Name="CarType" />
            <asp:Parameter Name="TicketLineNo" />
            <asp:Parameter Name="ActuakRun" />
            <asp:Parameter Name="ActualKM" />
            <asp:Parameter Name="Mins" />
            <asp:Parameter Name="AllowanceKM" />
            <asp:Parameter Name="IsExtra" />
            <asp:Parameter Name="Normal" />
            <asp:Parameter Name="FeedKM" />
            <asp:Parameter Name="FeedAMT1" />
            <asp:Parameter Name="FeedAMT2" />
            <asp:Parameter Name="BSDBSPKM" />
            <asp:Parameter Name="Toll" />
            <asp:Parameter Name="ApprovedDepNo" />
            <asp:Parameter Name="Operation" />
            <asp:Parameter Name="FixFile" />
            <asp:Parameter Name="IsFeed" />
            <asp:Parameter Name="CustNo" />
            <asp:Parameter Name="Amount" />
            <asp:Parameter Name="UnTaxAMT" />
            <asp:Parameter Name="DAteE" />
            <asp:Parameter Name="FixedCarCount" />
            <asp:Parameter Name="MainGovOffice" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridLines_List" Name="LinesNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="LineName" />
            <asp:Parameter Name="CallName" />
            <asp:Parameter Name="LinesNoM" />
            <asp:Parameter Name="LinesGovNo" />
            <asp:Parameter Name="LinesGOVExtNo" />
            <asp:Parameter Name="CarType" />
            <asp:Parameter Name="TicketLineNo" />
            <asp:Parameter Name="ActualRun" />
            <asp:Parameter Name="HighWay" />
            <asp:Parameter Name="IsExtra" />
            <asp:Parameter Name="BSDBSPKM" />
            <asp:Parameter Name="Operation" />
            <asp:Parameter Name="IsFeed" />
            <asp:Parameter Name="ActualKM" />
            <asp:Parameter Name="Mins" />
            <asp:Parameter Name="AllowanceKM" />
            <asp:Parameter Name="Normal" />
            <asp:Parameter Name="FeedKM" />
            <asp:Parameter Name="FeedAMT2" />
            <asp:Parameter Name="FeedAMT1" />
            <asp:Parameter Name="ApprovedTimes" />
            <asp:Parameter Name="Toll" />
            <asp:Parameter Name="ApprovedDepNo" />
            <asp:Parameter Name="FixFile" />
            <asp:Parameter Name="CustNo" />
            <asp:Parameter Name="Amount" />
            <asp:Parameter Name="UnTaxAMT" />
            <asp:Parameter Name="DateE" />
            <asp:Parameter Name="FixedCarCount" />
            <asp:Parameter Name="MainGovOffice" />
            <asp:Parameter Name="LinesNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
</asp:Content>
