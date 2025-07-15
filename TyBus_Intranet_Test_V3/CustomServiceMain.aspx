<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CustomServiceMain.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CustomServiceMain" %>

<asp:Content ID="CustomServiceMainForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">0800客服專線紀錄表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                    <asp:Label ID="lbSearchTitle" runat="server" CssClass="titleText-Red" Text="資　　　　料　　　　查　　　　詢" Width="100%" />
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
                    <asp:Button ID="bbPrint_Choised" runat="server" CssClass="button-Black" Text="選取列印" OnClick="bbPrint_Choised_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPrint_All" runat="server" CssClass="button-Black" Text="全部列印" OnClick="bbPrint_All_Click" Width="90%" />
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
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
        <asp:GridView ID="gridDataList" runat="server" AllowPaging="True" AutoGenerateColumns="False" DataKeyNames="ServiceNo" DataSourceID="dsServiceDataList"
            CellPadding="4" GridLines="None" ForeColor="#333333" PageSize="5">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:CheckBox ID="cbPrint_Detail" runat="server" OnCheckedChanged="cbPrint_Detail_CheckedChanged" />
                    </ItemTemplate>
                </asp:TemplateField>
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
                <asp:Label ID="lbEmptyData" runat="server" CssClass="text-Left-Red" Text="無符合條件資料" />
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
        <asp:SqlDataSource ID="dsServiceDataList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
            ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
            SelectCommand="SELECT ServiceNo, BuildDate, BuildTime, BuildMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuildMan)) AS BuildMan_C, ServiceType, (SELECT TypeText FROM CustomServiceType WHERE (TypeStep = '1') AND (TypeLevel1 = a.ServiceType)) AS ServiceType_C, ServiceTypeB, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE (TypeStep = '2') AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB)) AS ServiceTypeB_C, ServiceTypeC, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE (TypeStep = '3') AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB) AND (TypeLevel3 = a.ServiceTypeC)) AS ServiceTypeC_C, LinesNo, Car_ID, Driver, DriverName, BoardTime, BoardStation, GetoffTime, GetoffStation, ServiceNote, IsReplied, IsPending, AssignDate, AssignMan, (SELECT NAME FROM EMPLOYEE AS Employee_2 WHERE (EMPNO = a.AssignMan)) AS AssignMan_C, IsClosed, CloseDate, CloseMan, (SELECT NAME FROM EMPLOYEE AS Employee_1 WHERE (EMPNO = a.CloseMan)) AS CloseMan_C, CaseSource, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = CAST('客服專線記錄表' AS char(16)) + CAST('fmCustomService' AS char(16)) + 'CaseSource') AND (CLASSNO = a.CaseSource)) AS CaseSource_C, ServiceDate, CASE WHEN isnull(IsTrue , '') = '' THEN '未查證' WHEN isnull(IsTrue , '') = 'V' THEN '查證屬實' WHEN isnull(IsTrue , '') = 'X' THEN '查無實據' END AS IsTrue_C FROM CustomService AS a WHERE (1 &lt;&gt; 1) ORDER BY BuildDate DESC"></asp:SqlDataSource>
        <asp:FormView ID="fvServiceDataView" runat="server" DataKeyNames="ServiceNo" DataSourceID="dsServiceDataDetail" Width="100%"
            OnDataBound="fvServiceDataView_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CssClass="button-Black" OnClick="bbOK_Edit_Click" Text="更新" Width="120px" UseSubmitBehavior="False" OnClientClick="this.disabled=true;" />
                <asp:Button ID="bbCancel_Edit" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceNo_Edit" runat="server" CssClass="text-Right-Blue" Text="客訴單號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eServiceNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceNo") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseSource_Edit" runat="server" CssClass="text-Right-Blue" Text="案件來源：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlCaseSource_Edit" runat="server" AutoPostBack="True" CssClass="text-Left-Black"
                                        DataSourceID="sdsCaseSource_Edit" DataTextField="ClassTxt" DataValueField="ClassNo"
                                        OnSelectedIndexChanged="ddlCaseSource_Edit_SelectedIndexChanged" Width="90%">
                                    </asp:DropDownList>
                                    <asp:Label ID="eCaseSource_Edit" runat="server" Text='<%# Eval("CaseSource") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceDate_Edit" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eServiceDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildMan_Edit" runat="server" CssClass="text-Right-Blue" Text="開單人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuildMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="35%" />
                                    <asp:Label ID="eBuildManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildDate_Edit" runat="server" CssClass="text-Right-Blue" Text="開單時間：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuildDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate", "{0:yyyy/MM/dd}") %>' Width="55%" />
                                    <asp:Label ID="eBuildTime_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildTime") %>' Width="35%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceType_Edit" runat="server" CssClass="text-Right-Blue" Text="反映事項：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:DropDownList ID="ddlServiceType_Edit" runat="server" CssClass="text-Left-Black" Width="90%"
                                        DataSourceID="dsServiceTypeLevel1" DataTextField="TypeText" DataValueField="TypeNo" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlServiceType_Edit_SelectedIndexChanged" />
                                    <asp:Label ID="eServiceType_Edit" runat="server" Text='<%# Eval("ServiceType") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:DropDownList ID="ddlServiceTypeB_Edit" runat="server" CssClass="text-Left-Black" Width="90%"
                                        AutoPostBack="True" OnSelectedIndexChanged="ddlServiceTypeB_Edit_SelectedIndexChanged" />
                                    <asp:Label ID="eServiceTypeB_Edit" runat="server" Text='<%# Eval("ServiceTypeB") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:DropDownList ID="ddlServiceTypeC_Edit" runat="server" CssClass="text-Left-Black" Width="90%"
                                        AutoPostBack="True" OnSelectedIndexChanged="ddlServiceTypeC_Edit_SelectedIndexChanged" />
                                    <asp:Label ID="eServiceTypeC_Edit" runat="server" Text='<%# Eval("ServiceTypeC") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="Title1 ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbServiceTitle_Edit" runat="server" Text="反應事項概述：" Width="90%" CssClass="text-Left-Blue" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNo_Edit" runat="server" CssClass="text-Right-Blue" Text="路線編號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLinesNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNo") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarID_Edit" runat="server" CssClass="text-Right-Blue" Text="車輛牌照：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAthorityDep_Edit" runat="server" Text="權責單位：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:DropDownList ID="ddlAthorityDepNo_Edit" runat="server" CssClass="text-Left-Black" Width="35%" AutoPostBack="True" DataSourceID="dsAthorityDepNo"
                                        DataTextField="NAME" DataValueField="DEPNO" OnSelectedIndexChanged="ddlAthorityDepNo_Edit_SelectedIndexChanged" />
                                    <asp:TextBox ID="eAthorityDepNote_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AthorityDepNote") %>' Enabled="false" Width="55%" />
                                    <br />
                                    <asp:Label ID="eAthorityDepNo_Edit" runat="server" Text='<%# Eval("AthorityDepNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDriverName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>'
                                        OnTextChanged="eDriverName_Edit_TextChanged" AutoPostBack="true" Width="90%" />
                                    <br />
                                    <asp:Label ID="eDriver_Edit" runat="server" Text='<%# Eval("Driver") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="Title1 ColWidth-8Col" colspan="8">
                                    <br />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNo2_Edit" runat="server" CssClass="text-Right-Blue" Text="路線編號2：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLinesNo2_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNo2") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarID2_Edit" runat="server" CssClass="text-Right-Blue" Text="車輛牌照2：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID2_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID2") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAthorityDep2_Edit" runat="server" Text="權責單位2：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:DropDownList ID="ddlAthorityDepNo2_Edit" runat="server" CssClass="text-Left-Black" Width="35%" AutoPostBack="True"
                                        DataSourceID="dsAthorityDepNo2" DataTextField="NAME" DataValueField="DEPNO"
                                        OnSelectedIndexChanged="ddlAthorityDepNo2_Edit_SelectedIndexChanged" />
                                    <asp:TextBox ID="eAthorityDepNote2_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AthorityDepNote2") %>' Enabled="false" Width="55%" />
                                    <br />
                                    <asp:Label ID="eAthorityDepNo2_Edit" runat="server" Text='<%# Eval("AthorityDepNo2") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver2_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛員2：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDriverName2_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName2") %>'
                                        OnTextChanged="eDriverName2_Edit_TextChanged" AutoPostBack="true" Width="90%" />
                                    <br />
                                    <asp:Label ID="eDriver2_Edit" runat="server" Text='<%# Eval("Driver2") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="Title1 ColWidth-8Col" colspan="8">
                                    <br />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBoardTime_Edit" runat="server" CssClass="text-Right-Blue" Text="上車時間：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBoardTime_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BoardTime") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbGetoffTime_Edit" runat="server" CssClass="text-Right-Blue" Text="下車時間：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eGetoffTime_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GetoffTime") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBoardStation_Edit" runat="server" CssClass="text-Right-Blue" Text="上車站牌：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBoardStation_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BoardStation") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbGetoffStation_Edit" runat="server" CssClass="text-Right-Blue" Text="下車站牌：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eGetoffStation_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GetoffStation") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_High ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceNote_Edit" runat="server" CssClass="text-Right-Blue" Text="情況概述：" Width="90%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eServiceNote_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceNote") %>' TextMode="MultiLine" Width="90%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="Title1 ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbCivicContectData_Edit" runat="server" Text="民眾連絡方式：" Width="90%" CssClass="text-Left-Blue" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eIsNoContect_Edit" runat="server" Checked='<%# Eval("IsNoContect") %>' Text="未留連絡資訊" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCivicName_Edit" runat="server" Text="姓名：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCivicName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicName") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCivicTelNo_Edit" runat="server" Text="電話號碼：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eCivicTelNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicTelNo") %>' Width="55%" />
                                    <asp:TextBox ID="eCivicTelExtNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicTelExtNo") %>' Width="35%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCivicCellPhone_Edit" runat="server" Text="行動電話：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCivicCellPhone_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicCellPhone") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCivicAddress_Edit" runat="server" Text="地址：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eCivicAddress_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicAddress") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCivicEMail_Edit" runat="server" Text="eMail：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eCivicEMail_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicEMail") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="Title1 ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbResult_Edit" runat="server" Text="處理方式：" Width="90%" CssClass="text-Left-Blue" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="eIsReplied_Edit" runat="server" Checked='<%# Eval("IsReplied") %>' Text="已回覆民眾" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" rowspan="2">
                                    <asp:Label ID="lbRemark_Edit" runat="server" Text="回覆概述：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-8Col" colspan="5" rowspan="2">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" Text='<%# Eval("Remark") %>' CssClass="text-Left-Black" Width="90%" Height="95%" TextMode="MultiLine" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:DropDownList ID="ddlIsTrue_Edit" runat="server" CssClass="text-Left-Black"
                                        AutoPostBack="true" OnSelectedIndexChanged="ddlIsTrue_Edit_SelectedIndexChanged" Width="90%" />
                                    <asp:Label ID="eIsTrue_Edit" runat="server" Text='<%# Eval("IsTrue") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="eIsPending_Edit" runat="server" Checked='<%# Eval("IsPending") %>' Text="分發待查" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignDate_Edit" runat="server" CssClass="text-Right-Blue" Text="受理日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignMan_Edit" runat="server" Text="受理人：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eAssignMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' Width="35%"
                                        OnTextChanged="eAssignMan_Edit_TextChanged" />
                                    <asp:Label ID="eAssignName_Edit" runat="server" CssClass="text-Left-Black" Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="eIsClosed_Edit" runat="server" Checked='<%# Eval("IsClosed") %>' Enabled="false" Text="已結案" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCloseDate_Edit" runat="server" CssClass="text-Right-Blue" Text="結案日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCloseDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CloseDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCloseMan_Edit" runat="server" Text="結案人：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:Label ID="eCloseMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CloseMan") %>' Width="35%" />
                                    <asp:Label ID="eCloseManName_Edit" runat="server" CssClass="text-Left-Black" Width="55%" />
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
            <HeaderTemplate>
                <asp:Label ID="lbHeader" runat="server" CssClass="titleText-Blue" Text="桃園客運督導課 0800 客服專線紀錄表" Width="90%"></asp:Label>
            </HeaderTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CssClass="button-Black" OnClick="bbOK_INS_Click" Text="確定" Width="120px" OnClientClick="this.disabled=true;" UseSubmitBehavior="False" />
                <asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceNo_Insert" runat="server" CssClass="text-Right-Blue" Text="客訴單號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eServiceNo_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceNo") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseSource_Insert" runat="server" CssClass="text-Right-Blue" Text="案件來源：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlCaseSource_Insert" runat="server" AutoPostBack="True" CssClass="text-Left-Black"
                                        DataSourceID="sdsCaseSource_Insert" DataTextField="CLASSTXT" DataValueField="CLASSNO"
                                        OnSelectedIndexChanged="ddlCaseSource_Insert_SelectedIndexChanged" Width="90%">
                                    </asp:DropDownList>
                                    <asp:Label ID="eCaseSource_Insert" runat="server" Text='<%# Eval("CaseSource") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceDate_Insert" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eServiceDate_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildMan_Insert" runat="server" CssClass="text-Right-Blue" Text="開單人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuildMan_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="35%" />
                                    <asp:Label ID="eBuildManName_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildDate_Insert" runat="server" CssClass="text-Right-Blue" Text="開單時間：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eBuildDate_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate", "{0:yyyy/MM/dd}") %>' Width="55%" />
                                    <asp:TextBox ID="eBuildTime_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildTime") %>' Width="35%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceType_Insert" runat="server" CssClass="text-Right-Blue" Text="反映事項：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:DropDownList ID="ddlServiceType_Insert" runat="server" CssClass="text-Left-Black" Width="90%"
                                        DataSourceID="dsServiceTypeLevel1" DataTextField="TypeText" DataValueField="TypeNo" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlServiceType_Insert_SelectedIndexChanged" />
                                    <asp:Label ID="eServiceType_Insert" runat="server" Text='<%# Eval("ServiceType") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:DropDownList ID="ddlServiceTypeB_Insert" runat="server" CssClass="text-Left-Black" Width="90%"
                                        AutoPostBack="True" OnSelectedIndexChanged="ddlServiceTypeB_Insert_SelectedIndexChanged" />
                                    <asp:Label ID="eServiceTypeB_Insert" runat="server" Text='<%# Eval("ServiceTypeB") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:DropDownList ID="ddlServiceTypeC_Insert" runat="server" CssClass="text-Left-Black" Width="90%"
                                        AutoPostBack="True" OnSelectedIndexChanged="ddlServiceTypeC_Insert_SelectedIndexChanged" />
                                    <asp:Label ID="eServiceTypeC_Insert" runat="server" Text='<%# Eval("ServiceTypeC") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="Title1 ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbServiceTitle_Insert" runat="server" Text="反應事項概述：" Width="90%" CssClass="text-Left-Blue" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNo_Insert" runat="server" CssClass="text-Right-Blue" Text="路線編號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLinesNo_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNo") %>' Width="90%" AutoPostBack="True" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarID_Insert" runat="server" CssClass="text-Right-Blue" Text="車輛牌照：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAthorityDep_Insert" runat="server" Text="權責單位：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:DropDownList ID="ddlAthorityDepNo_Insert" runat="server" CssClass="text-Left-Black" Width="35%"
                                        DataSourceID="dsAthorityDepNo" DataTextField="NAME" DataValueField="DEPNO"
                                        OnSelectedIndexChanged="ddlAthorityDepNo_Insert_SelectedIndexChanged" AutoPostBack="True" />
                                    <asp:TextBox ID="eAthorityDepNote_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AthorityDepNote") %>' Enabled="false" Width="55%" />
                                    <br />
                                    <asp:Label ID="eAthorityDepNo_Insert" runat="server" Text='<%# Eval("AthorityDepNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_Insert" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eDriverName_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>'
                                        OnTextChanged="eDriverName_Insert_TextChanged" AutoPostBack="true" Width="90%" />
                                    <br />
                                    <asp:Label ID="eDriver_Insert" runat="server" Text='<%# Eval("Driver") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="Title1 ColWidth-8Col" colspan="8">
                                    <br />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLinesNo2_Insert" runat="server" CssClass="text-Right-Blue" Text="路線編號2：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eLinesNo2_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNo2") %>' Width="90%" AutoPostBack="True" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarID2_Insert" runat="server" CssClass="text-Right-Blue" Text="車輛牌照2：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID2_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID2") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAthorityDep2_Insert" runat="server" Text="權責單位2：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:DropDownList ID="ddlAthorityDepNo2_Insert" runat="server" CssClass="text-Left-Black" Width="35%"
                                        DataSourceID="dsAthorityDepNo2" DataTextField="NAME" DataValueField="DEPNO" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlAthorityDepNo2_Insert_SelectedIndexChanged" />
                                    <asp:TextBox ID="eAthorityDepNote2_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AthorityDepNote2") %>' Enabled="false" Width="55%" />
                                    <br />
                                    <asp:Label ID="eAthorityDepNo2_Insert" runat="server" Text='<%# Eval("AthorityDepNo2") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver2_Insert" runat="server" CssClass="text-Right-Blue" Text="駕駛員2：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eDriverName2_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName2") %>'
                                        OnTextChanged="eDriverName2_Insert_TextChanged" AutoPostBack="true" Width="90%" />
                                    <br />
                                    <asp:Label ID="eDriver2_Insert" runat="server" Text='<%# Eval("Driver2") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="Title1 ColWidth-8Col" colspan="8">
                                    <br />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBoardTime_Insert" runat="server" CssClass="text-Right-Blue" Text="上車時間：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBoardTime_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BoardTime") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbGetoffTime_Insert" runat="server" CssClass="text-Right-Blue" Text="下車時間：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eGetoffTime_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GetoffTime") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBoardStation_Insert" runat="server" CssClass="text-Right-Blue" Text="上車站牌：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBoardStation_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BoardStation") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbGetoffStation_Insert" runat="server" CssClass="text-Right-Blue" Text="下車站牌：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eGetoffStation_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GetoffStation") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_High ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceNote_Insert" runat="server" CssClass="text-Right-Blue" Text="情況概述：" Width="90%" />
                                </td>
                                <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eServiceNote_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceNote") %>' TextMode="MultiLine" Width="90%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="Title1 ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbCivicContectData_Insert" runat="server" Text="民眾連絡方式：" Width="90%" CssClass="text-Left-Blue" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eIsNoContect_Insert" runat="server" Text="未留連絡資訊" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCivicName_Insert" runat="server" Text="姓名：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCivicName_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicName") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCivicTelNo_Insert" runat="server" Text="電話號碼：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eCivicTelNo_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicTelNo") %>' Width="55%" />
                                    <asp:TextBox ID="eCivicTelExtNo_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicTelExtNo") %>' Width="35%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCivicCellPhone_Insert" runat="server" Text="行動電話：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCivicCellPhone_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicCellPhone") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCivicAddress_Insert" runat="server" Text="地址：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eCivicAddress_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicAddress") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCivicEMail_Insert" runat="server" Text="eMail：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eCivicEMail_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicEMail") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="Title1 ColWidth-8Col" colspan="8">
                                    <asp:Label ID="lbResult_Insert" runat="server" Text="處理方式：" Width="90%" CssClass="text-Left-Blue" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="eIsReplied_Insert" runat="server" Text="已回覆民眾" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" rowspan="2">
                                    <asp:Label ID="lbRemark_Insert" runat="server" Text="回覆概述：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-8Col" colspan="5" rowspan="2">
                                    <asp:TextBox ID="eRemark_Insert" runat="server" Text='<%# Eval("Remark") %>' CssClass="text-Left-Black" Width="90%" Height="95%" TextMode="MultiLine" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:DropDownList ID="ddlIsTrue_Insert" runat="server" CssClass="text-Left-Black"
                                        AutoPostBack="true" OnSelectedIndexChanged="ddlIsTrue_Insert_SelectedIndexChanged" Width="90%" />
                                    <asp:Label ID="eIsTrue_Insert" runat="server" Text='<%# Eval("IsTrue") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="eIsPending_Insert" runat="server" Text="分發待查" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignDate_Insert" runat="server" CssClass="text-Right-Blue" Text="受理日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAssignDate_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAssignMan_Insert" runat="server" Text="受理人：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eAssignMan_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' Width="35%"
                                        OnTextChanged="eAssignMan_Insert_TextChanged" />
                                    <asp:Label ID="eAssignName_Insert" runat="server" CssClass="text-Left-Black" Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="eIsClosed_Insert" runat="server" Enabled="false" Text="已結案" CssClass="text-Left-Black" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCloseDate_Insert" runat="server" CssClass="text-Right-Blue" Text="結案日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCloseDate_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CloseDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCloseMan_Insert" runat="server" Text="結案人：" CssClass="text-Right-Blue" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:Label ID="eCloseMan_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CloseMan") %>' Width="35%" />
                                    <asp:Label ID="eCloseManName_Insert" runat="server" CssClass="text-Left-Black" Width="55%" />
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
                <asp:Button ID="bbNew_List" Text="新增" runat="server" CssClass="button-Black" CausesValidation="false" CommandName="New" Width="90px" />
                <asp:Button ID="bbEdit_List" Text="修改" runat="server" CssClass="button-Red" CausesValidation="false" CommandName="Edit" Width="90px" />
                <asp:Button ID="bbDelete" Text="刪除" runat="server" CssClass="button-Red" CausesValidation="false" OnClick="bbDelete_Click" Width="90px" />
                <asp:Button ID="bbClosed" Text="結案" runat="server" CssClass="button-Red" CausesValidation="false" OnClick="bbClosed_Click" Width="90px" />
                <asp:Button ID="bbReopen" Text="取消結案" runat="server" CssClass="button-Black" CausesValidation="false" OnClick="bbReopen_Click" Width="90px" />
                <asp:Button ID="bbPrint_Single" Text="單張列印" runat="server" CssClass="button-Black" CausesValidation="false" OnClick="bbPrint_Single_Click" Width="90px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbServiceNo_List" runat="server" CssClass="text-Right-Blue" Text="客訴單號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eServiceNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceNo") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseSource_List" runat="server" CssClass="text-Right-Blue" Text="案件來源：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseSource_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseSource_C") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbServiceDate_List" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eServiceDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildMan_List" runat="server" CssClass="text-Right-Blue" Text="開單人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuildMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="35%" />
                            <asp:Label ID="eBuildManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildManName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildDate_List" runat="server" CssClass="text-Right-Blue" Text="開單時間：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuildDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate", "{0:yyyy/MM/dd}") %>' Width="55%" />
                            <asp:Label ID="eBuildTime_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildTime") %>' Width="35%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbServiceType_List" runat="server" CssClass="text-Right-Blue" Text="反映事項：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eServiceType_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceType_C") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eServiceTypeB_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceTypeB_C") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="ServiceTypeCLabel" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceTypeC_C") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="Title1 ColWidth-8Col" colspan="8">
                            <asp:Label ID="lbServiceTitle_List" runat="server" Text="反映事項概述：" Width="90%" CssClass="text-Left-Blue" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLinesNo_List" runat="server" CssClass="text-Right-Blue" Text="路線編號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLinesNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNo") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCarID_List" runat="server" CssClass="text-Right-Blue" Text="車輛牌照：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCar_ID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAthorityDep_List" runat="server" Text="權責單位：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eAthorityDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AthorityDepNo") %>' Width="35%" />
                            <asp:Label ID="eAthorityDepNote_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AthorityDepNote") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriver_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDriver_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver") %>' Width="35%" />
                            <asp:Label ID="eDriverName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="Title1 ColWidth-8Col" colspan="8">
                            <br />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLinesNo2_List" runat="server" CssClass="text-Right-Blue" Text="路線編號2：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLinesNo2_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LinesNo2") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCarID2_List" runat="server" CssClass="text-Right-Blue" Text="車輛牌照2：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCar_ID2_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID2") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAthorityDep2_List" runat="server" Text="權責單位2：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eAthorityDepNo2_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AthorityDepNo2") %>' Width="35%" />
                            <asp:Label ID="eAthorityDepNote2_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AthorityDepNote2") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriver2_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員2：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDriver2_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver2") %>' Width="35%" />
                            <asp:Label ID="eDriverName2_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName2") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="Title1 ColWidth-8Col" colspan="8">
                            <br />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBoardTime_List" runat="server" CssClass="text-Right-Blue" Text="上車時間：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBoardTime_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BoardTime") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbGetoffTime_List" runat="server" CssClass="text-Right-Blue" Text="下車時間：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eGetoffTime_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GetoffTime") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBoardStation_List" runat="server" CssClass="text-Right-Blue" Text="上車站牌：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBoardStation_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BoardStation") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbGetoffStation_List" runat="server" CssClass="text-Right-Blue" Text="下車站牌：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eGetoffStation_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GetoffStation") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_High ColBorder ColWidth-8Col">
                            <asp:Label ID="lbServiceNote_List" runat="server" CssClass="text-Right-Blue" Text="情況概述：" Width="90%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eServiceNote_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceNote") %>' ReadOnly="true" TextMode="MultiLine" Width="90%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="Title1 ColWidth-8Col" colspan="8">
                            <asp:Label ID="lbCivicContectData_List" runat="server" Text="民眾連絡方式：" Width="90%" CssClass="text-Left-Blue" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eIsNoContect_List" runat="server" Checked='<%# Eval("IsNoContect") %>' Enabled="false" Text="未留連絡資訊" CssClass="text-Left-Black" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCivicName_List" runat="server" Text="姓名：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCivicName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicName") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCivicTelNo_List" runat="server" Text="電話號碼：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eCivicTelNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicTelNo") %>' Width="55%" />
                            <asp:Label ID="eCivicTelExtNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicTelExtNo") %>' Width="35%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCivicCellPhone_List" runat="server" Text="行動電話：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCivicCellPhone_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicCellPhone") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCivicAddress_List" runat="server" Text="地址：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eCivicAddress_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicAddress") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCivicEMail_List" runat="server" Text="eMail：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eCivicEMail_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CivicEMail") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="Title1 ColWidth-8Col" colspan="8">
                            <asp:Label ID="lbResult_List" runat="server" Text="處理方式：" Width="90%" CssClass="text-Left-Blue" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:CheckBox ID="eIsReplied_List" runat="server" Checked='<%# Eval("IsReplied") %>' Enabled="false" Text="已回覆民眾" CssClass="text-Left-Black" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" rowspan="2">
                            <asp:Label ID="lbRemark_List" runat="server" Text="回覆概述：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="MultiLine_Low ColBorder ColWidth-8Col" colspan="5" rowspan="2">
                            <asp:TextBox ID="eRemark_List" runat="server" Text='<%# Eval("Remark") %>' ReadOnly="true" CssClass="text-Left-Black" Width="90%" Height="95%" TextMode="MultiLine" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:DropDownList ID="ddlIsTrue_List" runat="server" CssClass="text-Left-Black" Enabled="false" Width="90%" />
                            <asp:Label ID="eIsTrue_List" runat="server" Text='<%# Eval("IsTrue") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:CheckBox ID="eIsPending_List" runat="server" Checked='<%# Eval("IsPending") %>' Enabled="false" CssClass="text-Left-Black" Text="分發待查" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignDate_List" runat="server" CssClass="text-Right-Blue" Text="受理日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAssignDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssignMan_List" runat="server" Text="受理人：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eAssignMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' Width="35%" />
                            <asp:Label ID="eAssignName_List" runat="server" CssClass="text-Left-Black" Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:CheckBox ID="eIsClosed_List" runat="server" Checked='<%# Eval("IsClosed") %>' Enabled="false" Text="已結案" CssClass="text-Left-Black" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCloseDate_List" runat="server" CssClass="text-Right-Blue" Text="結案日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCloseDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CloseDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCloseMan_List" runat="server" Text="結案人：" CssClass="text-Right-Blue" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eCloseMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CloseMan") %>' Width="35%" />
                            <asp:Label ID="eCloseManName_List" runat="server" CssClass="text-Left-Black" Width="55%" />
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
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" Text="新增" runat="server" CssClass="button-Black" CausesValidation="false" CommandName="New" Width="120px" />
            </EmptyDataTemplate>
        </asp:FormView>
        <asp:GridView ID="gridServiceDetaildataList" runat="server" AutoGenerateColumns="False" CellPadding="4"
            DataKeyNames="ServiceNoItem" DataSourceID="dsServiceDetailDataList" ForeColor="#333333" GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:BoundField DataField="ServiceNoItem" HeaderText="ServiceNoItem" ReadOnly="True" SortExpression="ServiceNoItem" Visible="False" />
                <asp:BoundField DataField="ServiceNo" HeaderText="ServiceNo" SortExpression="ServiceNo" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" ItemStyle-Width="60px" ItemStyle-CssClass="fixedWidth" />
                <asp:BoundField DataField="ContactDate" DataFormatString="{0:d}" HeaderText="連絡日期" SortExpression="ContactDate" />
                <asp:BoundField DataField="ContactPerson" HeaderText="處理人" SortExpression="ContactPerson" />
                <asp:BoundField DataField="Excutort" HeaderText="連絡對象" SortExpression="Excutort" />
                <asp:BoundField DataField="ContactNote" HeaderText="處理概況" SortExpression="ContactNote" />
                <asp:BoundField DataField="AssignDate" DataFormatString="{0:d}" HeaderText="轉交日期" SortExpression="AssignDate" />
                <asp:BoundField DataField="AssignedMan" HeaderText="轉交人" SortExpression="AssignedMan" />
                <asp:BoundField DataField="BuildDate" HeaderText="BuildDate" SortExpression="BuildDate" Visible="False" />
                <asp:BoundField DataField="BuildMan" HeaderText="BuildMan" SortExpression="BuildMan" Visible="False" />
                <asp:BoundField DataField="ModifyDate" DataFormatString="{0:d}" HeaderText="異動日期" SortExpression="ModifyDate" />
                <asp:BoundField DataField="ModifyMan" HeaderText="異動人" SortExpression="ModifyMan" />
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
        <asp:SqlDataSource ID="dsServiceDataDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
            ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
            SelectCommand="SELECT ServiceNo, BuildDate, BuildTime, BuildMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuildMan)) AS BuildManName, ServiceType, (SELECT TypeText FROM CustomServiceType WHERE (TypeStep = '1') AND (TypeLevel1 = a.ServiceType)) AS ServiceType_C, ServiceTypeB, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE (TypeStep = '2') AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB)) AS ServiceTypeB_C, ServiceTypeC, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE (TypeStep = '3') AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB) AND (TypeLevel3 = a.ServiceTypeC)) AS ServiceTypeC_C, LinesNo, Car_ID, Driver, DriverName, LinesNo2, Car_ID2, Driver2, DriverName2, BoardTime, BoardStation, GetoffTime, GetoffStation, ServiceNote, IsNoContect, CivicName, CivicTelNo, CivicTelExtNo, CivicCellPhone, CivicAddress, CivicEMail, AthorityDepNo, AthorityDepNote, AthorityDepNo2, AthorityDepNote2, IsReplied, Remark, IsPending, AssignDate, AssignMan, (SELECT NAME FROM EMPLOYEE AS Employee_2 WHERE (EMPNO = a.AssignMan)) AS AssignManName, IsClosed, CloseDate, CloseMan, (SELECT NAME FROM EMPLOYEE AS Employee_1 WHERE (EMPNO = a.CloseMan)) AS CloseManName, CaseSource, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = CAST('客服專線記錄表' AS char(16)) + CAST('fmCustomService' AS char(16)) + 'CaseSource') AND (CLASSNO = a.CaseSource)) AS CaseSource_C, ServiceDate, IsTrue FROM CustomService AS a WHERE (ServiceNo = @ServiceNo)"
            InsertCommand="INSERT INTO CustomService(ServiceNo, BuildDate, BuildTime, BuildMan, ServiceType, ServiceTypeB, ServiceTypeC, LinesNo, Car_ID, Driver, DriverName, AthorityDepNo, AthorityDepNote, LinesNo2, Car_ID2, Driver2, DriverName2, AthorityDepNo2, AthorityDepNote2, BoardTime, BoardStation, GetoffTime, GetoffStation, ServiceNote, IsNoContect, CivicName, CivicTelNo, CivicTelExtNo, CivicCellPhone, CivicAddress, CivicEMail, IsReplied, Remark, IsPending, AssignDate, AssignMan, CaseSource, ServiceDate, IsTrue) VALUES (@ServiceNo, @BuildDate, @BuildTime, @BuildMan, @ServiceType, @ServiceTypeB, @ServiceTypeC, @LinesNo, @Car_ID, @Driver, @DriverName, @AthorityDepNo, @AthorityDepNote, @LinesNo2, @Car_ID2, @Driver2, @DriverName2, @AthorityDepNo2, @AthorityDepNote2, @BoardTime, @BoardStation, @GetoffTime, @GetoffStation, @ServiceNote, @IsNoContect, @CivicName, @CivicTelNo, @CivicTelExtNo, @CivicCellPhone, @CivicAddress, @CivicEMail, @IsReplied, @Remark, @IsPending, @AssignDate, @AssignMan, @CaseSource, @ServiceDate, @IsTrue)"
            UpdateCommand="UPDATE CustomService SET ServiceType = @ServiceType, ServiceTypeB = @ServiceTypeB, ServiceTypeC = @ServiceTypeC, LinesNo = @LinesNo, Car_ID = @Car_ID, Driver = @Driver, DriverName = @DriverName, BoardTime = @BoardTime, BoardStation = @BoardStation, GetoffTime = @GetoffTime, GetoffStation = @GetoffStation, ServiceNote = @ServiceNote, IsNoContect = @IsNoContect, CivicName = @CivicName, CivicTelNo = @CivicTelNo, CivicCellPhone = @CivicCellPhone, CivicAddress = @CivicAddress, CivicEMail = @CivicEMail, AthorityDepNo = @AthorityDepNo, AthorityDepNote = @AthorityDepNote, IsReplied = @IsReplied, Remark = @Remark, IsPending = @IsPending, AssignDate = @AssignDate, AssignMan = @AssignMan, CivicTelExtNo = @CivicTelExtNo, LinesNo2 = @LinesNo2, Car_ID2 = @Car_ID2, Driver2 = @Driver2, DriverName2 = @DriverName2, AthorityDepNo2 = @AthorityDepNo2, AthorityDepNote2 = @AthorityDepNote2, CaseSource = @CaseSource, ServiceDate = @ServiceDate, IsTrue = @IsTrue WHERE (ServiceNo = @ServiceNo)"
            DeleteCommand="DELETE FROM CustomService WHERE (ServiceNo = @ServiceNo)"
            OnInserting="dsServiceDataDetail_Inserting"
            OnDeleted="dsServiceDataDetail_Deleted"
            OnInserted="dsServiceDataDetail_Inserted"
            OnUpdated="dsServiceDataDetail_Updated"
            OnUpdating="dsServiceDataDetail_Updating">
            <DeleteParameters>
                <asp:Parameter Name="ServiceNo" />
            </DeleteParameters>
            <InsertParameters>
                <asp:Parameter Name="ServiceNo" />
                <asp:Parameter Name="BuildDate" />
                <asp:Parameter Name="BuildTime" />
                <asp:Parameter Name="BuildMan" />
                <asp:Parameter Name="ServiceType" />
                <asp:Parameter Name="ServiceTypeB" />
                <asp:Parameter Name="ServiceTypeC" />
                <asp:Parameter Name="LinesNo" />
                <asp:Parameter Name="Car_ID" />
                <asp:Parameter Name="Driver" />
                <asp:Parameter Name="DriverName" />
                <asp:Parameter Name="AthorityDepNo" />
                <asp:Parameter Name="AthorityDepNote" />
                <asp:Parameter Name="LinesNo2" />
                <asp:Parameter Name="Car_ID2" />
                <asp:Parameter Name="Driver2" />
                <asp:Parameter Name="DriverName2" />
                <asp:Parameter Name="AthorityDepNo2" />
                <asp:Parameter Name="AthorityDepNote2" />
                <asp:Parameter Name="BoardTime" />
                <asp:Parameter Name="BoardStation" />
                <asp:Parameter Name="GetoffTime" />
                <asp:Parameter Name="GetoffStation" />
                <asp:Parameter Name="ServiceNote" />
                <asp:Parameter Name="IsNoContect" />
                <asp:Parameter Name="CivicName" />
                <asp:Parameter Name="CivicTelNo" />
                <asp:Parameter Name="CivicTelExtNo" />
                <asp:Parameter Name="CivicCellPhone" />
                <asp:Parameter Name="CivicAddress" />
                <asp:Parameter Name="CivicEMail" />
                <asp:Parameter Name="IsReplied" />
                <asp:Parameter Name="Remark" />
                <asp:Parameter Name="IsPending" />
                <asp:Parameter Name="AssignDate" />
                <asp:Parameter Name="AssignMan" />
                <asp:Parameter Name="CaseSource" />
                <asp:Parameter Name="ServiceDate" />
                <asp:Parameter Name="IsTrue" />
            </InsertParameters>
            <SelectParameters>
                <asp:ControlParameter ControlID="gridDataList" Name="ServiceNo" PropertyName="SelectedValue" Type="String" />
            </SelectParameters>
            <UpdateParameters>
                <asp:Parameter Name="ServiceType" />
                <asp:Parameter Name="ServiceTypeB" />
                <asp:Parameter Name="ServiceTypeC" />
                <asp:Parameter Name="LinesNo" />
                <asp:Parameter Name="Car_ID" />
                <asp:Parameter Name="Driver" />
                <asp:Parameter Name="DriverName" />
                <asp:Parameter Name="BoardTime" />
                <asp:Parameter Name="BoardStation" />
                <asp:Parameter Name="GetoffTime" />
                <asp:Parameter Name="GetoffStation" />
                <asp:Parameter Name="ServiceNote" />
                <asp:Parameter Name="IsNoContect" />
                <asp:Parameter Name="CivicName" />
                <asp:Parameter Name="CivicTelNo" />
                <asp:Parameter Name="CivicCellPhone" />
                <asp:Parameter Name="CivicAddress" />
                <asp:Parameter Name="CivicEMail" />
                <asp:Parameter Name="AthorityDepNo" />
                <asp:Parameter Name="AthorityDepNote" />
                <asp:Parameter Name="IsReplied" />
                <asp:Parameter Name="Remark" />
                <asp:Parameter Name="IsPending" />
                <asp:Parameter Name="AssignDate" />
                <asp:Parameter Name="AssignMan" />
                <asp:Parameter Name="CivicTelExtNo" />
                <asp:Parameter Name="LinesNo2" />
                <asp:Parameter Name="Car_ID2" />
                <asp:Parameter Name="Driver2" />
                <asp:Parameter Name="DriverName2" />
                <asp:Parameter Name="AthorityDepNo2" />
                <asp:Parameter Name="AthorityDepNote2" />
                <asp:Parameter Name="CaseSource" />
                <asp:Parameter Name="ServiceDate" />
                <asp:Parameter Name="IsTrue" />
                <asp:Parameter Name="ServiceNo" />
            </UpdateParameters>
        </asp:SqlDataSource>
    </asp:Panel>

    <asp:SqlDataSource ID="dsServiceDetailDataList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT ServiceNoItem, ServiceNo, Items, ContactPerson, ContactNote, ContactDate, Excutort, AssignDate, AssignedMan, Remark, BuildDate, BuildMan, ModifyDate, ModifyMan FROM CustomServiceDetail WHERE (ServiceNo = @ServiceNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridDataList" Name="ServiceNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsServiceTypeLevel1" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar(3)) AS TypeNo, CAST(NULL AS varchar) AS TypeText UNION ALL SELECT LEFT (TypeNo, 3) AS TypeNo, TypeText FROM CustomServiceType WHERE (TypeStep = '1') ORDER BY TypeNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsAthorityDepNo" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar) AS DEPNO, CAST(NULL AS varchar) AS NAME UNION ALL SELECT DEPNO, NAME FROM DEPARTMENT WHERE (DEPNO BETWEEN '01' AND '32') AND (ISNULL(InPayReport, 'N') = 'Y') UNION ALL SELECT CAST('99' AS varchar) AS DEPNO, CAST('其他' AS varchar) AS NAME ORDER BY Depno"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsAthorityDepNo2" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar) AS DEPNO, CAST(NULL AS varchar) AS NAME UNION ALL SELECT DEPNO, NAME FROM DEPARTMENT WHERE (DEPNO BETWEEN '01' AND '32') AND (ISNULL(InPayReport, 'N') = 'Y') UNION ALL SELECT CAST('99' AS varchar) AS DEPNO, CAST('其他' AS varchar) AS NAME ORDER BY Depno"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsServiceType1_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar(3)) AS TypeNo, CAST(NULL AS varchar) AS TypeText UNION ALL SELECT LEFT (TypeNo, 3) AS TypeNo, TypeText FROM CustomServiceType WHERE (TypeStep = '1') ORDER BY TypeNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsServiceTypeB_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar(6)) AS TypeNo, CAST(NULL AS varchar) AS TypeText UNION ALL SELECT LEFT (TypeNo, 6) AS TypeNo, TypeText FROM CustomServiceType WHERE (LEFT (TypeNo, 3) = @TypeNo) AND (TypeStep = '2') ORDER BY TypeNo">
        <SelectParameters>
            <asp:ControlParameter ControlID="eServiceType_Search" Name="TypeNo" PropertyName="Text" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsServiceTypeC_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CAST(NULL AS varchar(9)) AS TypeNo, CAST(NULL AS varchar) AS TypeText UNION ALL SELECT TypeNo, TypeText FROM CUstomServiceType WHERE (LEFT (TypeNo, 6) = @TypeNo) AND (TypeStep = '3') ORDER BY TypeNo">
        <SelectParameters>
            <asp:ControlParameter ControlID="eServiceTypeB_Search" Name="TypeNo" PropertyName="Text" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsAthorityDep_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar) AS AthorityDepNo, CAST(NULL AS varchar) AS NoteText UNION ALL SELECT DISTINCT AthorityDepNo, AthorityDepNo + ':' + ISNULL(AthorityDepNote, '----') AS NoteText FROM customservice WHERE (ISNULL(AthorityDepNo, '') &lt;&gt; '') ORDER BY AthorityDepNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsBuildEmpNo_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar) AS BuildMan, CAST(NULL AS varchar) AS BuildName UNION ALL SELECT DISTINCT BuildMan, BuildMan + ':' + ISNULL((SELECT NAME FROM Employee WHERE (EMPNO = a.BuildMan)), '') AS BuiltName FROM CustomService AS a"></asp:SqlDataSource>

    <asp:SqlDataSource ID="sdsCaseSource_Insert" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = CAST('客服專線記錄表' AS char(16)) + CAST('fmCustomService' AS char(16)) + 'CaseSource') ORDER BY CLASSNO"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsCaseSource_Edit" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = CAST('客服專線記錄表' AS char(16)) + CAST('fmCustomService' AS char(16)) + 'CaseSource') ORDER BY CLASSNO"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsCaseSource_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = CAST('客服專線記錄表' AS char(16)) + CAST('fmCustomService' AS char(16)) + 'CaseSource') ORDER BY CLASSNO"></asp:SqlDataSource>

</asp:Content>
