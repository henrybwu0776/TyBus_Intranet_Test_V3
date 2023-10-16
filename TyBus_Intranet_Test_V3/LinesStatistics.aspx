<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="LinesStatistics.aspx.cs" Inherits="TyBus_Intranet_Test_V3.LinesStatistics" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="LinesStatisticsForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">各路線合格率統計</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:RadioButtonList ID="rbListType" runat="server" CssClass="text-Left-Black" RepeatColumns="3" Width="90%">
                        <asp:ListItem Value="01" Text="市區公車每月不合格天數" Selected="True" />
                        <asp:ListItem Value="02" Text="免費巴士合格率" />
                        <asp:ListItem Value="03" Text="國道公路合格率" />
                    </asp:RadioButtonList>
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbSourceDate" runat="server" CssClass="text-Right-Blue" Text="匯入資料日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eSourceDate" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:FileUpload ID="fuExcel" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCountDate_Search" runat="server" CssClass="text-Right-Blue" Text="統計日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eCountDate_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eCountDate_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbMinQualifiedRate" runat="server" CssClass="text-Right-Blue" Text="合格比例下限" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eMinQualifiedRate" runat="server" CssClass="text-Left-Black" Text="90" Width="80%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Right-Blue" Text="％" Width="10%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPreview" runat="server" CssClass="button-Black" Text="預覽報表" OnClick="bbPreview_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Red" Text="匯出 Excel" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbImport" runat="server" CssClass="button-Black" Text="匯入資料" OnClick="bbImport_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col" colspan="8">
                    <asp:Label ID="lbAlert" runat="server" CssClass="errorMessageText" Text="請先將從行車監控網站下載的 EXCEL 檔案的第一行取消合併儲存格並存檔後再進行轉入！" />
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
    <asp:Panel ID="plReport" runat="server" CssClass="PrintPanel">
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
            <LocalReport ReportPath="Report\LinesStatisticsP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsLinesStatistics" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        InsertCommand="INSERT INTO LinesStatistics(IndexNo, LinesNo, Station, ShiftDate, ShiftMode, LinesNote, ApprovedShift, QualifiedShift, UnqualifiedShift, QualifiedRate, IdentificationRate, OntimeRate, Remark, BuMan, BuDate, IdentificationShift, InsertIDCardShift, DriverInputShift, OnTimeTotalShift, OntimeShift) VALUES (@IndexNo, @LinesNo, @Station, @ShiftDate, @ShiftMode, @LinesNote, @ApprovedShift, @QualifiedShift, @UnqualifiedShift, @QualifiedRate, @IdentificationRate, @OntimeRate, @Remark, @BuMan, @BuDate, @IdentificationShift, @InsertIDCardShift, @DriverInputShift, @OnTimeTotalShift, @OntimeShift)"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" OnInserted="sdsLinesStatistics_Inserted">
        <InsertParameters>
            <asp:Parameter Name="IndexNo" />
            <asp:Parameter Name="LinesNo" />
            <asp:Parameter Name="Station" />
            <asp:Parameter Name="ShiftDate" />
            <asp:Parameter Name="ShiftMode" />
            <asp:Parameter Name="LinesNote" />
            <asp:Parameter Name="ApprovedShift" />
            <asp:Parameter Name="QualifiedShift" />
            <asp:Parameter Name="UnqualifiedShift" />
            <asp:Parameter Name="QualifiedRate" />
            <asp:Parameter Name="IdentificationRate" />
            <asp:Parameter Name="OntimeRate" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="BuDate" />
            <asp:Parameter Name="IdentificationShift" />
            <asp:Parameter Name="InsertIDCardShift" />
            <asp:Parameter Name="DriverInputShift" />
            <asp:Parameter Name="OnTimeTotalShift" />
            <asp:Parameter Name="OntimeShift" />
        </InsertParameters>
    </asp:SqlDataSource>
</asp:Content>
