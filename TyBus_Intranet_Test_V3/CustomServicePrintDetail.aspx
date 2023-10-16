<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CustomServicePrintDetail.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CustomServicePrintDetail" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="CustomServicePrintDetailForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">0800客服專線紀錄表列印</a>
    </div>
    <br />
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="100px" />
        <rsweb:ReportViewer ID="rvPrint" runat="server" BackColor="" ClientIDMode="AutoID" HighlightBackgroundColor="" InternalBorderColor="204, 204, 204" InternalBorderStyle="Solid" InternalBorderWidth="1px" LinkActiveColor="" LinkActiveHoverColor="" LinkDisabledColor="" PrimaryButtonBackgroundColor="" PrimaryButtonForegroundColor="" PrimaryButtonHoverBackgroundColor="" PrimaryButtonHoverForegroundColor="" SecondaryButtonBackgroundColor="" SecondaryButtonForegroundColor="" SecondaryButtonHoverBackgroundColor="" SecondaryButtonHoverForegroundColor="" SplitterBackColor="" ToolbarDividerColor="" ToolbarForegroundColor="" ToolbarForegroundDisabledColor="" ToolbarHoverBackgroundColor="" ToolbarHoverForegroundColor="" ToolBarItemBorderColor="" ToolBarItemBorderStyle="Solid" ToolBarItemBorderWidth="1px" ToolBarItemHoverBackColor="" ToolBarItemPressedBorderColor="51, 102, 153" ToolBarItemPressedBorderStyle="Solid" ToolBarItemPressedBorderWidth="1px" ToolBarItemPressedHoverBackColor="153, 187, 226" Height="600px" Width="100%">
            <LocalReport ReportPath="Report\CustomServiceDetail.rdlc">
                <DataSources>
                    <rsweb:ReportDataSource DataSourceId="sdsCustomServiceDetailP" Name="CustomServiceMain" />
                </DataSources>
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsCustomServiceDetailP" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT ServiceNo, BuildDate, BuildTime, BuildMan, (SELECT NAME FROM Employee WHERE (EMPNO = CustomService.BuildMan)) AS BuildManName, ServiceType, (SELECT TypeText FROM CustomServiceType WHERE (TypeStep = '1') AND (TypeLevel1 = CustomService.ServiceType)) AS ServiceType_C, ServiceTypeB, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE (TypeStep = '2') AND (TypeLevel1 = CustomService.ServiceType) AND (TypeLevel2 = CustomService.ServiceTypeB)) AS ServiceTypeB_C, ServiceTypeC, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE (TypeStep = '3') AND (TypeLevel1 = CustomService.ServiceType) AND (TypeLevel2 = CustomService.ServiceTypeB) AND (TypeLevel3 = CustomService.ServiceTypeC)) AS ServiceTypeC_C, LinesNo, Car_ID, Driver, DriverName, BoardTime, BoardStation, GetoffTime, GetoffStation, ServiceNote, IsNoContect, CivicName, CivicTelNo, CivicCellPhone, CivicAddress, CivicEMail, AthorityDepNo, AthorityDepNote, (SELECT NAME FROM Department WHERE (DEPNO = CustomService.AthorityDepNo)) AS AthorityDepName, IsReplied, Remark, IsPending, AssignDate, AssignMan, (SELECT NAME FROM Employee AS Employee_2 WHERE (EMPNO = CustomService.AssignMan)) AS AssignManName, IsClosed, CloseDate, CloseMan, (SELECT NAME FROM Employee AS Employee_1 WHERE (EMPNO = CustomService.CloseMan)) AS CloseManName FROM CustomService"></asp:SqlDataSource>
</asp:Content>
