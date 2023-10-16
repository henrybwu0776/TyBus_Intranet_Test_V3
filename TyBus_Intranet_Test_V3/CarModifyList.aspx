<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CarModifyList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CarModifyList" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="CarModifyListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">車輛調動通知單</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbActionDate_Search" runat="server" CssClass="text-Right-Blue" Text="發文日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eActionDate_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPreview" runat="server" CssClass="button-Black" Text="預覽報表" OnClick="bbPreview_Click" Width="90%" />
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
            <LocalReport ReportPath="Report\CarModifyListP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsCarModifyListP" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT e.CAR_ID AS CarID, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '車輛資料作業    Car_infoA       CAR_CLASS') AND (CLASSNO = a.Car_Class)) AS Car_Class_C, CONVERT (nvarchar(10), a.ProdDate, 111) AS ProdDate, a.sitqty AS SiteQty, CONVERT (nvarchar(10), e.TRANDATE, 111) AS TranDate, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.depno_O)) AS DepNo_O_C, CASE WHEN e.Tran_Type &lt;&gt; '1' THEN NULL ELSE (SELECT [Name] FROM Department WHERE DepNo = e.DepNo) END AS DepNo_C, e.Tran_Type, (CASE WHEN e.Tran_Type = '1' AND e.DepNo_O = e.DepNo AND (SELECT COUNT(CarTranNo) FROM Car_InfoE WHERE Car_ID = e.Car_ID) = 1 THEN '新進車輛' WHEN e.Tran_Type = '1' AND (SELECT COUNT(CarTranNo) FROM Car_InfoE WHERE Car_ID = e.Car_ID) &lt;&gt; 1 THEN '調站' ELSE (SELECT ClassTxt FROM DBDICB WHERE FKey = '車輛異動        Car_infoE       TRANTYPE' AND ClassNo = e.Tran_Type) END) AS Tran_Type_C, e.REMARK FROM Car_infoE AS e LEFT OUTER JOIN Car_infoA AS a ON a.Car_ID = e.CAR_ID WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
