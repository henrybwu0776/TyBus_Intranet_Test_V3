<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CustomServiceHistory.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CustomServiceHistory" %>

<asp:Content ID="CustomServiceHistoryForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">0800客服專線記錄表異動記錄</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" GroupingText="查詢條件" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbModifyMan_Search" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:DropDownList ID="ddlModifyMan_Search" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True"
                        DataSourceID="dsModifyEmpNo_Search" DataTextField="ModifyName" DataValueField="ModifyMan"
                        OnSelectedIndexChanged="ddlModifyMan_Search_SelectedIndexChanged" />
                    <br />
                    <asp:Label ID="eModifyMan_Search" runat="server" CssClass="text-Left-Black" Visible="false" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbModifyDate_Search" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eStartDate_Modify_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_Modify_Search" runat="server" CssClass="titleText-S-Blue" Text="～" Width="10%" />
                    <asp:TextBox ID="eEndDate_Modify_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColWidth-6Col"></td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbBuildMan_Search" runat="server" CssClass="text-Right-Blue" Text="開單人員：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:DropDownList ID="ddlBuildMan_Searech" runat="server" CssClass="text-Left-Black" Width="90%"
                        AutoPostBack="true" DataSourceID="dsBuildEmpNo_Search" DataTextField="BuildName" DataValueField="BuildMan"
                        OnSelectedIndexChanged="ddlBuildMan_Searech_SelectedIndexChanged" />
                    <br />
                    <asp:Label ID="eBuildMan_Search" runat="server" CssClass="text-Left-Black" Visible="false" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="Label1" runat="server" CssClass="text-Right-Blue" Text="開單日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eStartDate_Build_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_Build_Search" runat="server" CssClass="titleText-S-Blue" Text="～" Width="10%" />
                    <asp:TextBox ID="eEndDate_Build_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColWidth-6Col"></td>
            </tr>
            <tr>
                <td class="ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="搜尋" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
                </td>
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plDetail" runat="server" GroupingText="查詢結果" CssClass="ShowPanel">
        <asp:FormView ID="fvServiceDataView" runat="server" DataKeyNames="HistoryIndex" DataSourceID="sdsHistoryDetail" Width="100%" OnDataBound="fvServiceDataView_DataBound">
            <ItemTemplate>
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbHistoryIndex" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eServiceDate" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                            <br />
                            <asp:Label ID="eHistoryIndex" runat="server" CssClass="text-Left-Black" Text='<%# Eval("HistoryIndex") %>' Width="90%" Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbServiceNo" runat="server" CssClass="text-Right-Blue" Text="原單號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eServiceNo" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceNo") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyType" runat="server" CssClass="text-Right-Blue" Text="異動種類：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eModifyType" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyType_C") %>' Width="25%" />
                            <asp:Label ID="eModifyDate" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="30%" />
                            <asp:Label ID="eModfyMan" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="40%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseSource" runat="server" CssClass="text-Right-Blue" Text="案件來源：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseSource" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseSource_C") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildDate" runat="server" CssClass="text-Right-Blue" Text="開單日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuildDate" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate", "{0:yyyy/MM/dd}") %>' Width="45%" />
                            <asp:Label ID="eBuildTime" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildTime") %>' Width="45%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildMan" runat="server" CssClass="text-Right-Blue" Text="開單人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuildManID" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="35%" />
                            <asp:Label ID="eBuildManName" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildManName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColWidth-8Col"></td>
                        <td class="ColHeight ColWidth-8Col"></td>
                    </tr>
                    <tr>
                        <td class="Title1 ColBorder ColWidth-8Col" colspan="8">
                            <asp:Label ID="lbServiceTitle" runat="server" Text="反映事項概述：" Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbServiceType_1" runat="server" CssClass="text-Right-Blue" Text="反映類別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eServiceType_1" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Type1_Name") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbServiceType_B" runat="server" CssClass="text-Right-Blue" Text="客訴類別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eServiceType_B" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Type2_Name") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbServiceType_C" runat="server" CssClass="text-Right-Blue" Text="客訴事項：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eServiceType_C" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Type3_Name") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLinesNo" runat="server" CssClass="text-Right-Blue" Text="路線代號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLinesNo" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNo") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCar_ID" runat="server" CssClass="text-Right-Blue" Text="牌照號碼：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCar_ID" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriverName" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eDriver" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver") %>' Width="35%" />
                            <asp:Label ID="eDriverName" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="60" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLinesNo2" runat="server" CssClass="text-Right-Blue" Text="路線代號2：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLinesNo2" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNo2") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCar_ID2" runat="server" CssClass="text-Right-Blue" Text="牌照號碼2：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCar_ID2" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID2") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriverName2" runat="server" CssClass="text-Right-Blue" Text="駕駛員2：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eDriver2" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver2") %>' Width="35%" />
                            <asp:Label ID="eDriverName2" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName2") %>' Width="60" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBoardTime" runat="server" CssClass="text-Right-Blue" Text="上車時間：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBoardTime" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BoardTime") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBoardStation" runat="server" CssClass="text-Right-Blue" Text="上車站牌：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBoardStation" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BoardStation") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbGetoffTime" runat="server" CssClass="text-Right-Blue" Text="下車時間：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eGetoffTime" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GetoffTime") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbGetoffStation" runat="server" CssClass="text-Right-Blue" Text="下車站牌：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eGetoffStation" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GetoffStation") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColBorder ColWidth-8Col">
                            <asp:Label ID="lbServiceNote" runat="server" CssClass="text-Right-Blue" Text="情況概述：" Width="90%" />
                        </td>
                        <td class="MultiLine_Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eServiceNote" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" ReadOnly="true" Text='<%# Eval("ServiceNote") %>' Width="98%" Height="98%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="Title1 ColBorder ColWidth-8Col" colspan="8">
                            <asp:Label ID="lbCivicContectData" runat="server" Text="民眾連絡方式：" Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:CheckBox ID="cbIsNoContect" runat="server" Checked='<%# Eval("IsNoContect") %>' Enabled="false" Text="未留連絡資料" CssClass="text-Left-Black" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCivicName" runat="server" Text="姓名：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCivicName" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicName") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCivicTelNo" runat="server" Text="電話號碼：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCivicTelNo" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicTelNo") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCivicCellPhone" runat="server" Text="行動電話：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCivicCellPhone" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicCellPhone") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCivicAddress" runat="server" Text="地址：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="4">
                            <asp:Label ID="eCivicAddress" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicAddress") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCivicEMail" runat="server" Text="eMail：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eCivicEMail" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicEMail") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="Title1 ColBorder ColWidth-8Col" colspan="8">
                            <asp:Label ID="lbResult" runat="server" Text="處理方式：" Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:CheckBox ID="eIsReplied" runat="server" Checked='<%# Eval("IsReplied") %>' Enabled="false" Text="已回覆民眾" CssClass="text-Left-Black" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="2">
                            <asp:Label ID="lbRemark" runat="server" Text="回覆概述：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="MultiLine_Low ColBorder ColWidth-8Col" colspan="5" rowspan="2">
                            <asp:TextBox ID="eRemark" runat="server" Text='<%# Eval("Remark") %>' ReadOnly="true" CssClass="text-Left-Black" Width="90%" Height="95%" TextMode="MultiLine" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:DropDownList ID="ddlIsTrue" runat="server" CssClass="text-Left-Black" Width="100%" DataSourceID="dsIsTrue_List" DataTextField="ClassTxt" DataValueField="ClassNo" />
                        </td>
                        <asp:TextBox ID="eIsTrue" runat="server" Text='<%# Eval("IsTrue") %>' ReadOnly="true" Visible="false" />
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:CheckBox ID="eIsPending" runat="server" Checked='<%# Eval("IsPending") %>' Enabled="false" CssClass="text-Left-Black" Text="分發待查" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignDate" runat="server" CssClass="text-Right-Blue" Text="受理日期 (西元年)：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAssignDate" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignMan" runat="server" Text="受理人：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eAssignMan" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="35%" />
                            <asp:Label ID="eAssignName" runat="server" CssClass="text-Left-Black" Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:CheckBox ID="eIsClosed" runat="server" Checked='<%# Eval("IsClosed") %>' Enabled="false" Text="已結案" CssClass="text-Left-Black" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCloseDate" runat="server" CssClass="text-Right-Blue" Text="結案日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCloseDate" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CloseDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCloseMan" runat="server" Text="結案人：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eCloseMan" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CloseManName") %>' Width="35%" />
                            <asp:Label ID="eCloseManName" runat="server" CssClass="text-Left-Black" Width="55%" />
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
        <asp:GridView ID="gridServiceHistoryList" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" CellPadding="3" DataKeyNames="HistoryIndex" DataSourceID="sdsServiceHistory" GridLines="Horizontal" Width="100%">
            <AlternatingRowStyle BackColor="#F7F7F7" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="HistoryIndex" HeaderText="序號" InsertVisible="False" ReadOnly="True" SortExpression="HistoryIndex" />
                <asp:BoundField DataField="ServiceNo" HeaderText="原單號" SortExpression="ServiceNo" />
                <asp:BoundField DataField="CaseSource_C" HeaderText="案件來源" />
                <asp:BoundField DataField="ServiceDate" DataFormatString="{0:d}" HeaderText="發生日期" SortExpression="ServiceDate" />
                <asp:BoundField DataField="BuildDate" DataFormatString="{0:d}" HeaderText="開單日期" SortExpression="BuildDate" />
                <asp:BoundField DataField="BuildTime" HeaderText="開單時間" SortExpression="BuildTime" />
                <asp:BoundField DataField="BuildMan" HeaderText="BuildMan" SortExpression="BuildMan" Visible="False" />
                <asp:BoundField DataField="BuildManName" HeaderText="開單人" ReadOnly="True" SortExpression="BuildManName" />
                <asp:BoundField DataField="ServiceType" HeaderText="ServiceType" SortExpression="ServiceType" Visible="False" />
                <asp:BoundField DataField="Type1_Name" HeaderText="反映類別" ReadOnly="True" SortExpression="Type1_Name" />
                <asp:BoundField DataField="ServiceTypeB" HeaderText="ServiceTypeB" SortExpression="ServiceTypeB" Visible="False" />
                <asp:BoundField DataField="Type2_Name" HeaderText="客訴類別" ReadOnly="True" SortExpression="Type2_Name" />
                <asp:BoundField DataField="ServiceTypeC" HeaderText="ServiceTypeC" SortExpression="ServiceTypeC" Visible="False" />
                <asp:BoundField DataField="Type3_Name" HeaderText="客訴事項" ReadOnly="True" SortExpression="Type3_Name" />
                <asp:BoundField DataField="LinesNo" HeaderText="路線編號" SortExpression="LinesNo" />
                <asp:BoundField DataField="Car_ID" HeaderText="牌照號碼" SortExpression="Car_ID" />
                <asp:BoundField DataField="Driver" HeaderText="Driver" SortExpression="Driver" Visible="False" />
                <asp:BoundField DataField="Driver_Name" HeaderText="駕駛員姓名" ReadOnly="True" SortExpression="Driver_Name" />
            </Columns>
            <EmptyDataTemplate>
                <asp:Label ID="lbEmptyData" runat="server" CssClass="text-Left-Blue" Text="無符合條件資料" />
            </EmptyDataTemplate>
            <FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
            <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#F7F7F7" />
            <PagerStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" HorizontalAlign="Right" />
            <RowStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="#F7F7F7" />
            <SortedAscendingCellStyle BackColor="#F4F4FD" />
            <SortedAscendingHeaderStyle BackColor="#5A4C9D" />
            <SortedDescendingCellStyle BackColor="#D8D8F0" />
            <SortedDescendingHeaderStyle BackColor="#3E3277" />
        </asp:GridView>
    </asp:Panel>

    <asp:SqlDataSource ID="sdsServiceHistory" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT HistoryIndex, ServiceNo, BuildDate, BuildTime, BuildMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuildMan)) AS BuildManName, ServiceType, (SELECT TypeText FROM CustomServiceType WHERE (TypeStep = 1) AND (TypeLevel1 = a.ServiceType)) AS Type1_Name, ServiceTypeB, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE (TypeStep = 2) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB)) AS Type2_Name, ServiceTypeC, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE (TypeStep = 3) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB) AND (TypeLevel3 = a.ServiceTypeC)) AS Type3_Name, LinesNo, Car_ID, Driver, (SELECT NAME FROM EMPLOYEE AS Employee_1 WHERE (EMPNO = a.Driver)) AS Driver_Name, LinesNo2, Car_ID2, Driver2, (SELECT NAME FROM EMPLOYEE AS Employee_1 WHERE (EMPNO = a.Driver2)) AS Driver_Name2, CaseSource, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = CAST('客服專線記錄表' AS char(16)) + CAST('fmCustomService' AS char(16)) + 'CaseSource') AND (CLASSNO = a.CaseSource)) AS CaseSource_C, ServiceDate, IsTrue FROM CustomServiceHistory AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsHistoryDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT HistoryIndex, ServiceNo, BuildDate, BuildTime, BuildMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuildMan)) AS BuildManName, ServiceType, (SELECT TypeText FROM CustomServiceType WHERE (TypeStep = 1) AND (TypeLevel1 = a.ServiceType)) AS Type1_Name, ServiceTypeB, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE (TypeStep = 2) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB)) AS Type2_Name, ServiceTypeC, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE (TypeStep = 3) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB) AND (TypeLevel3 = a.ServiceTypeC)) AS Type3_Name, LinesNo, Car_ID, Driver, DriverName, LinesNo2, Car_ID2, Driver2, DriverName2, BoardTime, BoardStation, GetoffTime, GetoffStation, ServiceNote, IsNoContect, CivicName, CivicTelNo, CivicCellPhone, CivicAddress, CivicEMail, AthorityDepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.AthorityDepNo)) AS AthorityDepName, AthorityDepNote, AthorityDepNo2, (SELECT NAME FROM DEPARTMENT AS Department_1 WHERE (DEPNO = a.AthorityDepNo2)) AS AthorityDepName2, AthorityDepNote2, IsReplied, Remark, IsPending, AssignDate, AssignMan, (SELECT NAME FROM EMPLOYEE AS EMployee_3 WHERE (EMPNO = a.AssignMan)) AS AssignManName, IsClosed, CloseDate, CloseMan, (SELECT NAME FROM EMPLOYEE AS EMployee_2 WHERE (EMPNO = a.CloseMan)) AS CloseManName, ModifyType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = CAST('客訴單歷史' AS char(16)) + CAST('fmServiceHistory' AS char(16)) + 'ModifyType') AND (CLASSNO = a.ModifyType)) AS ModifyType_C, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMployee_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName, CaseSource, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (FKEY = CAST('客服專線記錄表' AS char(16)) + CAST('fmCustomService' AS char(16)) + 'CaseSource') AND (CLASSNO = a.CaseSource)) AS CaseSource_C, ServiceDate, IsTrue FROM CustomServiceHistory AS a WHERE (HistoryIndex = @HistoryIndex)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridServiceHistoryList" Name="HistoryIndex" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsIsTrue_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT CAST(NULL AS varchar) AS ClassNo, CAST(NULL AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '客服專線記錄表  fmCustomService IsTrue') AND (CLASSNO &lt;&gt; 'IT02') ORDER BY ClassNo" />
    <asp:SqlDataSource ID="dsModifyEmpNo_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CAST(NULL AS varchar) AS ModifyMan, CAST(NULL AS varchar) AS ModifyName UNION ALL SELECT DISTINCT ModifyMan, ModifyMan + ':' + ISNULL((SELECT NAME FROM Employee WHERE (EMPNO = a.ModifyMan)), '') AS ModifyName FROM CustomServiceHistory AS a"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsBuildEmpNo_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CAST(NULL AS varchar) AS BuildMan, CAST(NULL AS varchar) AS BuildName UNION ALL SELECT DISTINCT BuildMan, BuildMan + ':' + ISNULL((SELECT NAME FROM Employee WHERE (EMPNO = a.BuildMan)), '') AS BuiltName FROM CustomServiceHistory AS a"></asp:SqlDataSource>
</asp:Content>
