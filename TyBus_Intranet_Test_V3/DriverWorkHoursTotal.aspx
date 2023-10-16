<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DriverWorkHoursTotal.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DriverWorkHoursTotal" %>

<asp:Content ID="DriverWorkHoursTotalForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">駕駛員工時累計查詢</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" GroupingText="查詢條件" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbCalYM_Search" runat="server" CssClass="text-Right-Blue" Text="查詢年月：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eCalYear_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbCalYear_Search" runat="server" CssClass="text-Left-Black" Text=" 年" Width="5%" />
                    <asp:TextBox ID="eCalMonth_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbCalMonth_Search" runat="server" CssClass="text-Left-Black" Text=" 月" Width="5%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbDriver_Search" runat="server" CssClass="text-Right-Blue" Text="駕駛員工號：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eDriverNo_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="eDriverName_Search" runat="server" CssClass="text-Left-Black" Width="50%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbStopDate_Search" runat="server" CssClass="text-Right-Blue" Text="截止日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eStopDate_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbClear" runat="server" CssClass="button-Red" Text="清除" OnClick="bbClear_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight_Low ColWidth-7Col">
                    <asp:Label ID="lbRecordCount" runat="server" Text="0" Visible="false" />
                </td>
                <td class="ColHeight_Low ColWidth-7Col" />
                <td class="ColHeight_Low ColWidth-7Col" />
                <td class="ColHeight_Low ColWidth-7Col" />
                <td class="ColHeight_Low ColWidth-7Col" />
                <td class="ColHeight_Low ColWidth-7Col" />
                <td class="ColHeight_Low ColWidth-7Col" />
                <td class="ColHeight_Low ColWidth-7Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" GroupingText="查詢結果" CssClass="ShowPanel">
        <asp:FormView ID="fvDriverWorkHoursTotalData" runat="server" Width="100%" DataSourceID="sdsDriverWorkHouyrsTotalData" OnDataBound="fvDriverWorkHoursTotalData_DataBound">
            <EmptyDataTemplate>
                <asp:Label ID="lbEmptyText" runat="server" CssClass="Title1" Text="查無符合條件的資料" Width="90%" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriver_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eDriverNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="30%" />
                            <asp:Label ID="eDriverName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBirthday_List" runat="server" CssClass="text-Right-Blue" Text="出生日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBirthday_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Birthday", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAssumeday_List" runat="server" CssClass="text-Right-Blue" Text="到職日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAssumeday_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Assumeday", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-8Col" colspan="8">
                            <asp:Label ID="lbAreaSplit_1" runat="server" CssClass="Title1" Text="員工工時統計" Width="100%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbHARM_List" runat="server" CssClass="text-Right-Blue" Text="應上時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eHARM_List" runat="server" CssClass="textBold" Text='<%# Eval("HARM") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbOther_List" runat="server" CssClass="text-Right-Blue" Text="月時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eOther_List" runat="server" CssClass="textBold" Text='<%# Eval("Other") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStatus_List" runat="server" CssClass="text-Right-Blue" Text="特殊狀態：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eHasOverTime_List" runat="server" CssClass="textRed" Text="● 有超時" Width="30%" Visible="false" />
                            <asp:Label ID="eWorkOverSevenDay_List" runat="server" CssClass="textRed" Text="▲ 超過七日未休" Width="30%" Visible="false" />
                            <asp:Label ID="eRetire_List" runat="server" CssClass="textRed" Text="★ 屆退人員" Width="30%" Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriverHR_List" runat="server" CssClass="text-Right-Blue" Text="行車工時：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDriverHR_List" runat="server" CssClass="textBold" Text='<%# Eval("WorkHR_T") %>' Width="30%" />
                            <asp:Label ID="lbSplit_1" runat="server" CssClass="textBold" Text="：" Width="5%" />
                            <asp:Label ID="eDriverMin_List" runat="server" CssClass="textBold" Text='<%# Eval("WorkMin") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLeave_List" runat="server" CssClass="text-Right-Blue" Text="請假工時：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLeave_List" runat="server" CssClass="textBold" Text='<%# Eval("Leave") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbWorkHR_List" runat="server" CssClass="text-Right-Blue" Text="累計工時：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eWorkHR_List" runat="server" CssClass="textBold" Text='<%# Eval("WorkHR") %>' Width="30%" />
                            <asp:Label ID="lbSplit_2" runat="server" CssClass="textBold" Text="：" Width="5%" />
                            <asp:Label ID="eWorkMin_List" runat="server" CssClass="textBold" Text='<%# Eval("WorkMin") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRent_List" runat="server" CssClass="text-Right-Blue" Text="租車趟次：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eRent_List" runat="server" CssClass="textBold" Text='<%# Eval("Rent") %>' />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbOverTime_List" runat="server" CssClass="text-Right-Blue" Text="休假加班：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eOverTime_List" runat="server" CssClass="textBold" Text='<%# Eval("OverTime") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbOverTime133_List" runat="server" CssClass="text-Right-Blue" Text="1.33加班：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eOverTime133_List" runat="server" CssClass="textBold" Text='<%# Eval("OverTime133") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbOverTime166_List" runat="server" CssClass="text-Right-Blue" Text="1.66加班：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eOverTime166_List" runat="server" CssClass="textBold" Text='<%# Eval("OverTime166") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbQK_List" runat="server" CssClass="text-Right-Blue" Text="已休天數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eQK_List" runat="server" CssClass="textBold" Text='<%# Eval("QK") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriverSum_List" runat="server" CssClass="text-Right-Blue" Text="個人累計 ( % )：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDriverSum_List" runat="server" CssClass="textBold" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStanderSum_List" runat="server" CssClass="text-Right-Blue" Text="累計標準 ( % )：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStanderSum_List" runat="server" CssClass="textBold" Width="90%" />
                        </td>
                        <td class="ColHeight ColWidth-8Col" colspan="4">
                            <asp:Label ID="eNotoffDay_List" runat="server" Text='<%# Eval("NotoffDay") %>' Visible="false" />
                            <asp:Label ID="eOT_List" runat="server" Text='<%# Eval("OT") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight_Low ColWidth-8Col" />
                        <td class="ColHeight_Low ColWidth-8Col" />
                        <td class="ColHeight_Low ColWidth-8Col" />
                        <td class="ColHeight_Low ColWidth-8Col" />
                        <td class="ColHeight_Low ColWidth-8Col" />
                        <td class="ColHeight_Low ColWidth-8Col" />
                        <td class="ColHeight_Low ColWidth-8Col" />
                        <td class="ColHeight_Low ColWidth-8Col" />
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsDriverWorkHouyrsTotalData" runat="server"
        ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar(12)) AS DepNo, CAST(NULL AS varchar(30)) AS DepName, CAST(NULL AS varchar(12)) AS EmpNo, CAST(NULL AS varchar(64)) AS EmpName, CAST(NULL AS DateTime) AS Birthday, CAST(NULL AS DateTime) AS Assumeday, CAST(0 AS int) AS HARM, CAST(0 AS int) AS Other, CAST(0 AS int) AS NotoffDay, CAST(0 AS float) AS OverTime, CAST(0 AS float) AS OverTime133, CAST(0 AS float) AS OverTime166, CAST(0 AS float) AS WorkHR, CAST(0 AS int) AS WorkMin, CAST(0 AS float) AS Rent, CAST(0 AS int) AS QK, CAST(0 AS float) AS Leave, CAST(0 AS int) AS OT"
        OnSelected="sdsDriverWorkHouyrsTotalData_Selected"></asp:SqlDataSource>
</asp:Content>
