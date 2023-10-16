<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CustomServiceDetail.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CustomServiceDetail" %>

<asp:Content ID="CustomServiceDetailForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">0800客服專線處理歷程紀錄表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                    <asp:Label ID="lbSearchTitle" runat="server" CssClass="titleText-Blue" Text="資　　　　料　　　　查　　　　詢" Width="100%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuildMan_Search" runat="server" CssClass="text-Right-Blue" Text="開單人員：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlBuildMan_Search" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True"
                        DataSourceID="dsBuildEmpNo_Search" DataTextField="BuildName" DataValueField="BuildMan"
                        OnSelectedIndexChanged="ddlBuildMan_Search_SelectedIndexChanged" />
                    <br />
                    <asp:Label ID="eBuildMan_Searech" runat="server" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbServiceType_Search" runat="server" CssClass="text-Right-Blue" Text="反映事項：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlServiceType_Search" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True"
                        DataSourceID="dsServiceType1_Search" DataTextField="TypeText" DataValueField="TypeNo"
                        OnSelectedIndexChanged="ddlServiceType_Search_SelectedIndexChanged" />
                    <br />
                    <asp:Label ID="eServiceType_Search" runat="server" Width="90%" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlServiceTypeB_Search" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True"
                        DataSourceID="dsServiceTypeB_Search" DataTextField="TypeText" DataValueField="TypeNo"
                        OnSelectedIndexChanged="ddlServiceTypeB_Search_SelectedIndexChanged" />
                    <br />
                    <asp:Label ID="eServiceTypeB_Search" runat="server" Width="90%" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlServiceTypeC_Search" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True"
                        DataSourceID="dsServiceTypeC_Search" DataTextField="TypeText" DataValueField="TypeNo"
                        OnSelectedIndexChanged="ddlServiceTypeC_Search_SelectedIndexChanged" />
                    <br />
                    <asp:Label ID="eServiceTypeC_Search" runat="server" Width="90%" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbLinesNo_Search" runat="server" CssClass="text-Right-Blue" Text="路線代號：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eLinesNo_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuildDate_Search" runat="server" CssClass="text-Right-Blue" Text="開單日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eBuildDate_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_Search_1" runat="server" CssClass="titleText-S-Blue" Text="～" Width="5%" />
                    <asp:TextBox ID="eBuildDate_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbServiceDate_Search" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eServiceDate_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_Search_2" runat="server" CssClass="titleText-S-Blue" Text="～" Width="5%" />
                    <asp:TextBox ID="eServiceDate_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCarID_Search" runat="server" CssClass="text-Right-Blue" Text="牌照號碼：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCarID_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDriver_Search" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDriver_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCaseSource_Search" runat="server" CssClass="text-Right-Blue" Text="資料來源" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:DropDownList ID="ddlCaseSource_Search" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True"
                        DataSourceID="sdsCaseSource_Search" DataTextField="ClassTxt" DataValueField="ClassNo"
                        OnSelectedIndexChanged="ddlCaseSource_Search_SelectedIndexChanged" />
                    <br />
                    <asp:Label ID="eCaseSource_Search" runat="server" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbAthorityDep_Search" runat="server" CssClass="text-Right-Blue" Text="權責單位" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlAthorityDep_Search" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True"
                        DataSourceID="dsAthorityDep_Search" DataTextField="NoteText" DataValueField="AthorityDepNo"
                        OnSelectedIndexChanged="ddlAthorityDep_Search_SelectedIndexChanged" />
                    <br />
                    <asp:Label ID="eAthorityDep_Search" runat="server" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCivicName_Search" runat="server" CssClass="text-Right-Blue" Text="客訴人姓名：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCivicName_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCivicTel_Search" runat="server" CssClass="text-Right-Blue" Text="客訴人電話：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCivicTel_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:RadioButtonList ID="rbIsClosed" runat="server" CssClass="text-Left-Black" Width="90%" RepeatColumns="3">
                        <asp:ListItem Value="0">全部紀錄</asp:ListItem>
                        <asp:ListItem Value="1">已結案</asp:ListItem>
                        <asp:ListItem Value="2" Selected="True">未結案</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbIsTrue_Search" runat="server" CssClass="text-Right-Blue" Text="查證情況：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlIsTrue_Search" runat="server" CssClass="text-Left-Black"
                        AutoPostBack="true" OnSelectedIndexChanged="ddlIsTrue_Search_SelectedIndexChanged" Width="90%">
                    </asp:DropDownList>
                    <br />
                    <asp:Label ID="eIsTrue_Search" runat="server" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbRequset" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbRequset_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
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
        <asp:GridView ID="gridDataList" runat="server" AllowPaging="True" AutoGenerateColumns="False" DataKeyNames="ServiceNo" DataSourceID="dsServiceDataList"
            CellPadding="4" GridLines="None" ForeColor="#333333" PageSize="5">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="CaseSource" HeaderText="CaseSource" SortExpression="CaseSource" Visible="false" />
                <asp:BoundField DataField="CaseSource_C" HeaderText="案件來源" SortExpression="CaseSource_C" ReadOnly="true" />
                <asp:BoundField DataField="ServiceNo" HeaderText="客訴單號" ReadOnly="True" SortExpression="ServiceNo" />
                <asp:BoundField DataField="BuildDate" DataFormatString="{0:D}" HeaderText="開單日期" SortExpression="BuildDate" />
                <asp:BoundField DataField="ServiceDate" DataFormatString="{0:D}" HeaderText="發生日期" SortExpression="ServiceDate" />
                <asp:BoundField DataField="BuildTime" HeaderText="開單時間" SortExpression="BuildTime" />
                <asp:BoundField DataField="BuildMan" HeaderText="開單人編號" SortExpression="BuildMan" Visible="False" />
                <asp:BoundField DataField="BuildMan_C" HeaderText="開單人" ReadOnly="True" SortExpression="BuildMan_C" />
                <asp:BoundField DataField="ServiceType" HeaderText="大類" SortExpression="ServiceType" Visible="False" />
                <asp:BoundField DataField="ServiceType_C" HeaderText="反映事項主類別" ReadOnly="True" SortExpression="ServiceType_C" />
                <asp:BoundField DataField="ServiceTypeB" HeaderText="次類別" SortExpression="ServiceTypeB" Visible="False" />
                <asp:BoundField DataField="ServiceTypeB_C" HeaderText="反映事項次類別" ReadOnly="True" SortExpression="ServiceTypeB_C" />
                <asp:BoundField DataField="ServiceTypeC" HeaderText="子類別" SortExpression="ServiceTypeC" Visible="False" />
                <asp:BoundField DataField="ServiceTypeC_C" HeaderText="反映事項子類別" ReadOnly="True" SortExpression="ServiceTypeC_C" />
                <asp:BoundField DataField="LinesNo" HeaderText="路線代號" SortExpression="LinesNo" />
                <asp:BoundField DataField="Car_ID" HeaderText="車牌號碼" SortExpression="Car_ID" />
                <asp:BoundField DataField="Driver" HeaderText="駕駛員編號" SortExpression="Driver" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員姓名" SortExpression="DriverName" />
                <asp:BoundField DataField="BoardTime" HeaderText="上車時間" SortExpression="BoardTime" />
                <asp:BoundField DataField="BoardStation" HeaderText="上車站牌" SortExpression="BoardStation" />
                <asp:BoundField DataField="GetoffTime" HeaderText="下車時間" SortExpression="GetoffTime" />
                <asp:BoundField DataField="GetoffStation" HeaderText="下車站牌" SortExpression="GetoffStation" />
                <asp:BoundField DataField="ServiceNote" HeaderText="反映事項" SortExpression="ServiceNote" />
                <asp:CheckBoxField DataField="IsReplied" HeaderText="已回覆客戶" SortExpression="IsReplied" />
                <asp:BoundField DataField="IsTrue_C" HeaderText="查證情況" SortExpression="IsTrue_C" />
                <asp:CheckBoxField DataField="IsPending" HeaderText="分發待查" SortExpression="IsPending" />
                <asp:BoundField DataField="AssignDate" DataFormatString="{0:D}" HeaderText="受理日期" SortExpression="AssignDate" />
                <asp:BoundField DataField="AssignMan" HeaderText="受理人編號" SortExpression="AssignMan" Visible="False" />
                <asp:BoundField DataField="AssignMan_C" HeaderText="受理人" ReadOnly="True" SortExpression="AssignMan_C" />
                <asp:CheckBoxField DataField="IsClosed" HeaderText="已結案" SortExpression="IsClosed" />
                <asp:BoundField DataField="CloseDate" HeaderText="CloseDate" SortExpression="CloseDate" Visible="False" />
                <asp:BoundField DataField="CloseMan" HeaderText="CloseMan" SortExpression="CloseMan" Visible="False" />
                <asp:BoundField DataField="CloseMan_C" HeaderText="CloseMan_C" ReadOnly="True" SortExpression="CloseMan_C" Visible="False" />
            </Columns>
            <EditRowStyle BackColor="#2461BF" />
            <EmptyDataTemplate>
                <asp:Label ID="lbEmptyData" runat="server" CssClass="text-Left-Blue" Text="無符合條件資料" />
            </EmptyDataTemplate>
            <FooterStyle BackColor="#507CD1" ForeColor="White" Font-Bold="True" />
            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#EFF3FB" />
            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#F5F7FB" />
            <SortedAscendingHeaderStyle BackColor="#6D95E1" />
            <SortedDescendingCellStyle BackColor="#E9EBEF" />
            <SortedDescendingHeaderStyle BackColor="#4870BE" />
        </asp:GridView>
        <asp:GridView ID="gridDetailDataList" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4"
            DataKeyNames="ServiceNoItem" DataSourceID="dsCustonServiceDetailList" ForeColor="#333333" GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="ServiceNoItem" HeaderText="ServiceNoItem" ReadOnly="True" SortExpression="ServiceNoItem" Visible="False" />
                <asp:BoundField DataField="ServiceNo" HeaderText="ServiceNo" SortExpression="ServiceNo" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                <asp:BoundField DataField="ContactDate" DataFormatString="{0:d}" HeaderText="連絡日期" SortExpression="ContactDate" />
                <asp:BoundField DataField="ContactPerson" HeaderText="連絡對象" SortExpression="ContactPerson" />
                <asp:BoundField DataField="Excutort" HeaderText="處理人" SortExpression="Excutort" Visible="False" />
                <asp:BoundField DataField="Excutort_C" HeaderText="處理人" ReadOnly="True" SortExpression="Excutort_C" />
                <asp:BoundField DataField="ContactNote" HeaderText="處理概述" SortExpression="ContactNote" />
                <asp:BoundField DataField="AssignDate" DataFormatString="{0:d}" HeaderText="轉交日期" SortExpression="AssignDate" />
                <asp:BoundField DataField="AssignedMan" HeaderText="AssignedMan" SortExpression="AssignedMan" Visible="False" />
                <asp:BoundField DataField="AssignedMan_C" HeaderText="轉交人" ReadOnly="True" SortExpression="AssignedMan_C" />
                <asp:BoundField DataField="BuildDate" DataFormatString="{0:d}" HeaderText="建檔日期" SortExpression="BuildDate" />
                <asp:BoundField DataField="BuildMan" HeaderText="BuildMan" SortExpression="BuildMan" Visible="False" />
                <asp:BoundField DataField="BuildMan_C" HeaderText="建檔人" ReadOnly="True" SortExpression="BuildMan_C" />
                <asp:BoundField DataField="ModifyDate" DataFormatString="{0:d}" HeaderText="異動日" SortExpression="ModifyDate" />
                <asp:BoundField DataField="ModifyMan" HeaderText="ModifyMan" SortExpression="ModifyMan" Visible="False" />
                <asp:BoundField DataField="ModifyMan_C" HeaderText="異動人" ReadOnly="True" SortExpression="ModifyMan_C" />
                <asp:BoundField DataField="Remark" HeaderText="異動說明" SortExpression="Remark" />
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
        <asp:FormView ID="fvCustomServiceDetail" runat="server" DataKeyNames="ServiceNoItem" DataSourceID="dsCustomServiceDetail" Width="100%"
            OnItemCreated="fvCustomServiceDetail_ItemCreated">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Update" runat="server" CssClass="button-Black" CausesValidation="True" CommandName="Update" Text="確定" />
                &nbsp;<asp:Button ID="bbCancel_Update" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceNo_Edit" runat="server" CssClass="text-Right-Blue" Text="客訴單號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eServiceNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceNo") %>' Width="90%" />
                                    <asp:Label ID="eServiceNoItem_Edit" runat="server" Text='<%# Eval("ServiceNoItem") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbItems_Edit" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="90%" />
                                    <br />
                                    <asp:Label ID="eBuildMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan_C") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbExcutort_Edit" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eExcutort_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eExcutort_Edit_TextChanged" Text='<%# Bind("Excutort") %>' Width="35%" />
                                    <asp:Label ID="eExcutort_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Excutort_C") %>' Width="55%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine_Low" rowspan="3">
                                    <asp:Label ID="lbContactNote_Edit" runat="server" CssClass="text-Right-Blue" Text="處理概述：" Width="90%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine_Low" colspan="3" rowspan="3">
                                    <asp:TextBox ID="eContactNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("ContactNote") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbContextDate_Edit" runat="server" CssClass="text-Right-Blue" Text="連絡日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eContactDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContactDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbContactPerson_Edit" runat="server" CssClass="text-Right-Blue" Text="連絡對象：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eContactPerson_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ContactPerson") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignDate_Edit" runat="server" CssClass="text-Right-Blue" Text="轉交日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Assigndate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignedMan_Edit" runat="server" CssClass="text-Right-Blue" Text="轉交人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignedMan_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignedMan_Edit_TextChanged" Text='<%# Bind("AssignedMan") %>' Width="90%" />
                                    <br />
                                    <asp:Label ID="eAssignedMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignedMan_C") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="90%" />
                                    <br />
                                    <asp:Label ID="eModifyMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="90%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine_Low" rowspan="2">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="異動說明：" Width="90%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine_Low" colspan="5" rowspan="2">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
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
            <EmptyDataTemplate>
                <asp:Button ID="bbInsert_Empty" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" />
            </EmptyDataTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CssClass="button-Black" CausesValidation="True" CommandName="Insert" Text="確定" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceNo_INS" runat="server" CssClass="text-Right-Blue" Text="客訴單號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eServiceNo_INS" runat="server" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbItems_INS" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eItems_INS" runat="server" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildDate_INS" runat="server" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildMan_INS" runat="server" CssClass="text-Left-Black" Width="90%" />
                                    <br />
                                    <asp:Label ID="eBuildMan_C_INS" runat="server" CssClass="text-Left-Black" Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbExcutort_INS" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eExcutort_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eExcutort_INS_TextChanged" Text='<%# Bind("Excutort") %>' Width="35%" />
                                    <asp:Label ID="eExcutort_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Excutort_C") %>' Width="55%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine_Low" rowspan="3">
                                    <asp:Label ID="lbContactNote_INS" runat="server" CssClass="text-Right-Blue" Text="處理概述：" Width="90%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine_Low" colspan="3" rowspan="3">
                                    <asp:TextBox ID="eContactNote_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("ContactNote") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbContextDate_INS" runat="server" CssClass="text-Right-Blue" Text="連絡日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eContactDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContactDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbContactPerson_INS" runat="server" CssClass="text-Right-Blue" Text="連絡對象：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eContactPerson_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ContactPerson") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignDate_INS" runat="server" CssClass="text-Right-Blue" Text="轉交日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Assigndate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignedMan_INS" runat="server" CssClass="text-Right-Blue" Text="轉交人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignedMan_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eAssignedMan_INS_TextChanged" Text='<%# Bind("AssignedMan") %>' Width="90%" />
                                    <br />
                                    <asp:Label ID="eAssignedMan_C_INS" runat="server" CssClass="text-Left-Black" Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="90%" />
                                    <br />
                                    <asp:Label ID="eModifyMan_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="90%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine_Low" rowspan="2">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="異動說明：" Width="90%" />
                                </td>
                                <td class="ColBorder ColWidth-8Col MultiLine_Low" colspan="5" rowspan="2">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Enabled="false" Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
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
            <ItemTemplate>
                <asp:Button ID="bbInsert_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="Edit" Text="修改" />
                &nbsp;<asp:Button ID="bbDelete_List" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Delete" Text="刪除" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbServiceNo_List" runat="server" CssClass="text-Right-Blue" Text="客訴單號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eServiceNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceNo") %>' Width="90%" />
                            <asp:Label ID="eServiceNoItem_List" runat="server" Text='<%# Eval("ServiceNoItem") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItems_List" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuildDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuildMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="90%" />
                            <br />
                            <asp:Label ID="eBuildMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan_C") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbExcutort_List" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eExcutort_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Excutort") %>' Width="35%" />
                            <asp:Label ID="eExcutort_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Excutort_C") %>' Width="55%" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine_Low" rowspan="3">
                            <asp:Label ID="lbContactNote_List" runat="server" CssClass="text-Right-Blue" Text="處理概述：" Width="90%" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine_Low" colspan="3" rowspan="3">
                            <asp:TextBox ID="eContactNote_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("ContactNote") %>' Enabled="false" Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbContextDate_List" runat="server" CssClass="text-Right-Blue" Text="連絡日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eContactDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContactDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbContactPerson_List" runat="server" CssClass="text-Right-Blue" Text="連絡對象：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eContactPerson_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ContactPerson") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignDate_List" runat="server" CssClass="text-Right-Blue" Text="轉交日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAssignDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Assigndate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignedMan_List" runat="server" CssClass="text-Right-Blue" Text="轉交人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAssignedMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignedMan") %>' Width="90%" />
                            <br />
                            <asp:Label ID="eAssignedMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignedMan_C") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="90%" />
                            <br />
                            <asp:Label ID="eModifyMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="90%" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine_Low" rowspan="2">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="異動說明：" Width="90%" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine_Low" colspan="5" rowspan="2">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Enabled="false" Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
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
    <asp:SqlDataSource ID="dsServiceDataList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT ServiceNo, BuildDate, BuildTime, BuildMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuildMan)) AS BuildMan_C, ServiceType, (SELECT TypeText FROM CustomServiceType WHERE (TypeStep = '1') AND (TypeLevel1 = a.ServiceType)) AS ServiceType_C, ServiceTypeB, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE (TypeStep = '2') AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB)) AS ServiceTypeB_C, ServiceTypeC, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE (TypeStep = '3') AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB) AND (TypeLevel3 = a.ServiceTypeC)) AS ServiceTypeC_C, LinesNo, Car_ID, Driver, DriverName, BoardTime, BoardStation, GetoffTime, GetoffStation, ServiceNote, IsReplied, IsPending, AssignDate, AssignMan, (SELECT NAME FROM EMPLOYEE AS Employee_2 WHERE (EMPNO = a.AssignMan)) AS AssignMan_C, IsClosed, CloseDate, CloseMan, (SELECT NAME FROM EMPLOYEE AS Employee_1 WHERE (EMPNO = a.CloseMan)) AS CloseMan_C, CaseSource, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = CAST('客服專線記錄表' AS char(16)) + CAST('fmCustomService' AS char(16)) + 'CaseSource') AND (CLASSNO = a.CaseSource)) AS CaseSource_C, ServiceDate, CASE WHEN isnull(IsTrue , '') = '' THEN '未查證' WHEN isnull(IsTrue , '') = 'V' THEN '查證屬實' WHEN isnull(IsTrue , '') = 'X' THEN '查無實據' END AS IsTrue_C FROM CustomService AS a WHERE (1 &lt;&gt; 1) ORDER BY BuildDate DESC"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsCustonServiceDetailList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT ServiceNoItem, ServiceNo, Items, ContactPerson, ContactNote, ContactDate, Excutort, AssignDate, AssignedMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = A.AssignedMan)) AS AssignedMan_C, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (EMPNO = A.Excutort)) AS Excutort_C, Remark, BuildDate, BuildMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = A.BuildMan)) AS BuildMan_C, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = A.ModifyMan)) AS ModifyMan_C FROM CustomServiceDetail AS A WHERE (ServiceNo = @ServiceNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridDataList" Name="ServiceNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsCustomServiceDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        DeleteCommand="DELETE FROM CustomServiceDetail WHERE (ServiceNoItem = @ServiceNoItem)"
        InsertCommand="INSERT INTO CustomServiceDetail(ServiceNoItem, ServiceNo, Items, ContactPerson, ContactNote, ContactDate, Excutort, BuildDate, BuildMan, AssignDate, AssignedMan) VALUES (@ServiceNoItem, @ServiceNo, @Items, @ContactPerson, @ContactNote, @ContactDate, @Excutort, @BuildDate, @BuildMan, @AssignDate, @AssignedMan)"
        SelectCommand="SELECT ServiceNoItem, ServiceNo, Items, ContactPerson, ContactNote, ContactDate, Excutort, AssignDate, AssignedMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = A.AssignedMan)) AS AssignedMan_C, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (EMPNO = A.Excutort)) AS Excutort_C, Remark, BuildDate, BuildMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = A.BuildMan)) AS BuildMan_C, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = A.ModifyMan)) AS ModifyMan_C FROM CustomServiceDetail AS A WHERE (ServiceNoItem = @ServiceNoItem)"
        UpdateCommand="UPDATE CustomServiceDetail SET ContactPerson = @ContactPerson, ContactNote = @ContactNote, ContactDate = @ContactDate, Excutort = @Excutort, Remark = @Remark, ModifyDate = @ModifyDate, ModifyMan = @ModifyMan, AssignDate = @AssignDate, AssignedMan = @AssignedMan WHERE (ServiceNoItem = @ServiceNoItem)"
        OnDeleted="dsCustomServiceDetail_Deleted"
        OnInserted="dsCustomServiceDetail_Inserted"
        OnInserting="dsCustomServiceDetail_Inserting"
        OnUpdated="dsCustomServiceDetail_Updated"
        OnUpdating="dsCustomServiceDetail_Updating">
        <DeleteParameters>
            <asp:Parameter Name="ServiceNoItem" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="ServiceNoItem" />
            <asp:Parameter Name="ServiceNo" />
            <asp:Parameter Name="Items" />
            <asp:Parameter Name="ContactPerson" />
            <asp:Parameter Name="ContactNote" />
            <asp:Parameter Name="ContactDate" />
            <asp:Parameter Name="Excutort" />
            <asp:Parameter Name="BuildDate" />
            <asp:Parameter Name="BuildMan" />
            <asp:Parameter Name="AssignDate" />
            <asp:Parameter Name="AssignedMan" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridDetailDataList" Name="ServiceNoItem" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="ContactPerson" />
            <asp:Parameter Name="ContactNote" />
            <asp:Parameter Name="ContactDate" />
            <asp:Parameter Name="Excutort" />
            <asp:Parameter Name="ServiceNoItem" />
            <asp:Parameter Name="AssignDate" />
            <asp:Parameter Name="AssignedMan" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="ModifyMan" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsServiceType1_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar(3)) AS TypeNo, CAST(NULL AS varchar) AS TypeText UNION ALL SELECT LEFT (TypeNo, 3) AS TypeNo, TypeText FROM CustomServiceType WHERE (TypeStep = '1') ORDER BY TypeNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsServiceTypeB_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar(6)) AS TypeNo, CAST(NULL AS varchar) AS TypeText UNION ALL SELECT LEFT (TypeNo, 6) AS TypeNo, TypeText FROM CustomServiceType WHERE (LEFT (TypeNo, 3) = @TypeNo) AND (TypeStep = '2') ORDER BY TypeNo">
        <SelectParameters>
            <asp:ControlParameter ControlID="eServiceType_Search" Name="TypeNo" PropertyName="Text" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsServiceTypeC_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar(9)) AS TypeNo, CAST(NULL AS varchar) AS TypeText UNION ALL SELECT TypeNo, TypeText FROM CUstomServiceType WHERE (LEFT (TypeNo, 6) = @TypeNo) AND (TypeStep = '3') ORDER BY TypeNo">
        <SelectParameters>
            <asp:ControlParameter ControlID="eServiceTypeB_Search" Name="TypeNo" PropertyName="Text" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsAthorityDep_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar) AS AthorityDepNo, CAST(NULL AS varchar) AS NoteText UNION ALL SELECT DISTINCT AthorityDepNo, AthorityDepNo + ':' + ISNULL(AthorityDepNote, '----') AS NoteText FROM customservice WHERE (ISNULL(AthorityDepNo, '') &lt;&gt; '') ORDER BY AthorityDepNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsBuildEmpNo_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar) AS BuildMan, CAST(NULL AS varchar) AS BuildName UNION ALL SELECT DISTINCT BuildMan, BuildMan + ':' + ISNULL((SELECT NAME FROM Employee WHERE (EMPNO = a.BuildMan)), '') AS BuiltName FROM CustomService AS a"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsCaseSource_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = CAST('客服專線記錄表' AS char(16)) + CAST('fmCustomService' AS char(16)) + 'CaseSource') ORDER BY CLASSNO"></asp:SqlDataSource>
</asp:Content>
