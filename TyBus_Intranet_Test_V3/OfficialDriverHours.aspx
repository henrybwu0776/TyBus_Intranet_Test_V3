<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="OfficialDriverHours.aspx.cs" Inherits="TyBus_Intranet_Test_V3.OfficialDriverHours" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="OfficialDriverHoursForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">公務車行車憑證登記表</a>
    </div>
    <br />
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plInsert" runat="server" Width="45%" GroupingText="新增行車時數" Visible="false">
        <asp:UpdatePanel ID="upDataInsert" runat="server">
            <ContentTemplate>
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbDriver_Insert" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder" colspan="2">
                            <asp:Label ID="eDriver_Insert" runat="server" CssClass="text-Left-Black" Width="30%" />
                            <asp:Label ID="eDriverName_Insert" runat="server" CssClass="text-Left-Black" Width="45%" />
                            <asp:Label ID="eDriverDate_Insert" runat="server" Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbCalDate_Insert" runat="server" CssClass="text-Right-Blue" Text="日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="eCalDate_Insert" runat="server" CssClass="text-Left-Black" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:DropDownList ID="ddlWorkState_Insert" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True"
                                DataSourceID="sdsWorkState_Insert" DataTextField="ClassTXT" DataValueField="ClassNo" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbWorkTime_Insert" runat="server" CssClass="text-Right-Blue" Text="上班時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1" colspan="2">
                            <asp:TextBox ID="eWorkHours_Insert" runat="server" CssClass="text-Left-Black" Width="20%" />
                            <asp:Label ID="lbSpilt1_Insert" runat="server" CssClass="text-Left-Black" Text="：" Width="5%" />
                            <asp:TextBox ID="eWorkMins_Insert" runat="server" CssClass="text-Left-Black" Width="20%" />
                            <asp:Label ID="eWorkTimes_Insert" runat="server" Visible="False" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-High ColBorder ColWidth1">
                            <asp:Label ID="lbRemark_Insert" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                        </td>
                        <td class="MultiLine-High ColBorder ColWidth1" colspan="2">
                            <asp:TextBox ID="eRemark_Insert" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1" colspan="3">
                            <asp:Label ID="lbTitle_Insert" runat="server" CssClass="titleText-S-Blue" Text="計算結果：" Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbExtra100_Insert" runat="server" CssClass="text-Right-Blue" Text="加給日薪時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1" colspan="2">
                            <asp:Label ID="eExtra100_Insert" runat="server" CssClass="text-Right-Blue" Width="25%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbExtra133_Insert" runat="server" CssClass="text-Right-Blue" Text="1.33 加班時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1" colspan="2">
                            <asp:Label ID="eExtra133_Insert" runat="server" CssClass="text-Right-Blue" Width="25%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbExtra166_Insert" runat="server" CssClass="text-Right-Blue" Text="1.66 加班時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1" colspan="2">
                            <asp:Label ID="eExtra166_Insert" runat="server" CssClass="text-Right-Blue" Width="25%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbExtra266_Insert" runat="server" CssClass="text-Right-Blue" Text="2.66 加班時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1" colspan="2">
                            <asp:Label ID="eExtra266_Insert" runat="server" CssClass="text-Right-Blue" Width="25%" />
                        </td>
                    </tr>
                </table>
            </ContentTemplate>
        </asp:UpdatePanel>
        <asp:Button ID="bbCalculate_Insert" runat="server" CssClass="button-Red" Text="計算" OnClick="bbCalculate_Insert_Click" Width="120px" />
        <asp:Button ID="bbInsert" runat="server" CssClass="button-Black" Text="確定" OnClick="bbInsert_Click" Width="120px" />
        <asp:Button ID="bbCancel_Insert" runat="server" CssClass="button-Red" Text="取消" OnClick="bbCancel_Insert_Click" Width="120px" />
    </asp:Panel>
    <asp:Panel ID="plEdit" runat="server" Width="45%" GroupingText="修改行車時數" Visible="false">
        <asp:UpdatePanel ID="upDataEdit" runat="server">
            <ContentTemplate>
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbDriver_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder" colspan="2">
                            <asp:Label ID="eDriver_Edit" runat="server" CssClass="text-Left-Black" Width="30%" />
                            <asp:Label ID="eDriverName_Edit" runat="server" CssClass="text-Left-Black" Width="45%" />
                            <asp:Label ID="eDriverDate_Edit" runat="server" Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbCalDate_Edit" runat="server" CssClass="text-Right-Blue" Text="日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="eCalDate_Edit" runat="server" CssClass="text-Left-Black" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:DropDownList ID="ddlWorkState_Edit" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True"
                                DataSourceID="sdsWorkState_Edit" DataTextField="ClassTXT" DataValueField="ClassNo" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbWorkTime_Edit" runat="server" CssClass="text-Right-Blue" Text="上班時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1" colspan="2">
                            <asp:TextBox ID="eWorkHours_Edit" runat="server" CssClass="text-Left-Black" Width="20%" />
                            <asp:Label ID="lbSpilt1_Edit" runat="server" CssClass="text-Left-Black" Text="：" Width="5%" />
                            <asp:TextBox ID="eWorkMins_Edit" runat="server" CssClass="text-Left-Black" Width="20%" />
                            <asp:Label ID="eWorkTimes_Edit" runat="server" Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-High ColBorder ColWidth1">
                            <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                        </td>
                        <td class="MultiLine-High ColBorder ColWidth1" colspan="2">
                            <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1" colspan="3">
                            <asp:Label ID="lbTitle_Edit" runat="server" CssClass="titleText-S-Blue" Text="計算結果：" Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbExtra100_Edit" runat="server" CssClass="text-Right-Blue" Text="加給日薪時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1" colspan="2">
                            <asp:Label ID="eExtra100_Edit" runat="server" CssClass="text-Right-Blue" Width="25%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbExtra133_Edit" runat="server" CssClass="text-Right-Blue" Text="1.33 加班時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1" colspan="2">
                            <asp:Label ID="eExtra133_Edit" runat="server" CssClass="text-Right-Blue" Width="25%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbExtra166_Edit" runat="server" CssClass="text-Right-Blue" Text="1.66 加班時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1" colspan="2">
                            <asp:Label ID="eExtra166_Edit" runat="server" CssClass="text-Right-Blue" Width="25%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth1">
                            <asp:Label ID="lbExtra266_Edit" runat="server" CssClass="text-Right-Blue" Text="2.66 加班時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth1" colspan="2">
                            <asp:Label ID="eExtra266_Edit" runat="server" CssClass="text-Right-Blue" Width="25%" />
                        </td>
                    </tr>
                </table>
            </ContentTemplate>
        </asp:UpdatePanel>
        <asp:Button ID="bbCalculate_Edit" runat="server" CssClass="button-Red" Text="計算" OnClick="bbCalculate_Edit_Click" Width="120px" />
        <asp:Button ID="bbModify" runat="server" CssClass="button-Black" Text="確定" OnClick="bbModify_Click" Width="120px" />
        <asp:Button ID="bbCancel_Edit" runat="server" CssClass="button-Red" Text="取消" OnClick="bbCancel_Edit_Click" Width="120px" />
    </asp:Panel>
    <asp:Panel ID="plSearch" runat="server" GroupingText="行事曆" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDriver_Search" runat="server" CssClass="text-Right-Blue" Text="查詢駕駛員：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlDriver_Search" runat="server" CssClass="text-Left-Black" DataSourceID="sdsDriverList" AutoPostBack="true"
                        DataTextField="NAME" DataValueField="EMPNO" OnSelectedIndexChanged="ddlDriver_Search_SelectedIndexChanged" Width="90%" />
                    <asp:Label ID="eDriver_Search" runat="server" Visible="false" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Black" Text="列印時數表" OnClick="bbPrint_Click" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col" colspan="8">
                    <asp:Calendar ID="calOfficialDriverHours" runat="server" Width="100%"
                        OnSelectionChanged="calOfficialDriverHours_SelectionChanged"
                        OnDayRender="calOfficialDriverHours_DayRender" BackColor="#CCFFFF">
                        <DayStyle Font-Bold="True" Font-Size="14pt" />
                        <TitleStyle Font-Bold="True" Font-Size="20pt" ForeColor="#FFFF66" BackColor="#6666FF" />
                    </asp:Calendar>
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
    <asp:Panel ID="plPrint" runat="server" CssClass="text-Left-Blue" GroupingText="報表列印" Width="90%" Visible="False">
        <asp:Button ID="bbCloseReport" runat="server" CssClass="button-Red" Text="結束預覽" OnClick="bbCloseReport_Click" Width="120px" />
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
            <LocalReport ReportPath="Report\OfficialDriverHoursP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsOfficialDriverHoursDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        InsertCommand="INSERT INTO OfficialDriverHours(DriverDate, CalDate, Driver, WorkHours, WorkMins, TotalWorkHours, Extra100, Extra133, Extra166, Extra266, Remark, WorkState) VALUES (@DriverDate, @CalDate, @Driver, @WorkHours, @WorkMins, @TotalWorkHours, @Extra100, @Extra133, @Extra166, @Extra266, @Remark, @WorkState)"
        UpdateCommand="UPDATE OfficialDriverHours SET CalDate = @CalDate, Driver = @Driver, WorkHours = @WorkHours, WorkMins = @WorkMins, TotalWorkHours = @TotalWorkHours, Extra100 = @Extra100, Extra133 = @Extra133, Extra166 = @Extra166, Extra266 = @Extra266, Remark = @Remark, WorkState = @WorkState WHERE (DriverDate = @DriverDate)">
        <InsertParameters>
            <asp:Parameter Name="DriverDate" />
            <asp:Parameter Name="CalDate" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="WorkHours" />
            <asp:Parameter Name="WorkMins" />
            <asp:Parameter Name="TotalWorkHours" />
            <asp:Parameter Name="Extra100" />
            <asp:Parameter Name="Extra133" />
            <asp:Parameter Name="Extra166" />
            <asp:Parameter Name="Extra266" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="WorkState" />
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="CalDate" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="WorkHours" />
            <asp:Parameter Name="WorkMins" />
            <asp:Parameter Name="TotalWorkHours" />
            <asp:Parameter Name="Extra100" />
            <asp:Parameter Name="Extra133" />
            <asp:Parameter Name="Extra166" />
            <asp:Parameter Name="Extra266" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="DriverDate" />
            <asp:Parameter Name="WorkState" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsDriverList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS varchar) AS EmpNo, CAST(NULL AS Varchar) AS Name UNION ALL SELECT EMPNO, NAME FROM Employee WHERE (LEAVEDAY IS NULL) AND (IsOfficialDriver = 'V') ORDER BY EmpNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsWorkState_Insert" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '行車記錄單A     runsheeta       WORKSTATE') ORDER BY CLASSNO"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsWorkState_Edit" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '行車記錄單A     runsheeta       WORKSTATE') ORDER BY CLASSNO"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsDriverHoursPrint" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"></asp:SqlDataSource>
</asp:Content>
