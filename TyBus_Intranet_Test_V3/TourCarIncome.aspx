<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="TourCarIncome.aspx.cs" Inherits="TyBus_Intranet_Test_V3.TourCarIncome" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="TourCarIncomeForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">各站遊覽車營收成本表</a>
    </div>
    <br />
    <asp:Panel ID="plInputData" runat="server" GroupingText="資料輸入" CssClass="ShowPanel">
        <table id="tableInputData" class="TableSetting">
            <tr>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Black TableColAlign_Right" colspan="2">行車年月 (西元年 4 碼 + 月份 2 碼 ; 範例：201801)</td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Black" colspan="4">
                    <asp:TextBox ID="DriveYM" runat="server" CssClass="InputMargin" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Blue TableColAlign_Right">每車年強制險保費</td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Blue">
                    <asp:TextBox ID="CarYearInsurance" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Blue TableColAlign_Right">每車年第三責任險</td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Blue">
                    <asp:TextBox ID="CarThirdInsurance" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Blue TableColAlign_Right"></td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Blue"></td>
            </tr>
            <tr>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Blue TableColAlign_Right">駕駛員每人月勞保費</td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Blue">
                    <asp:TextBox ID="DriverLiMonth" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Blue TableColAlign_Right">駕駛員每人月健保費</td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Blue">
                    <asp:TextBox ID="DriverHiMonth" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Blue TableColAlign_Right">駕駛員獎金每人每月</td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Blue">
                    <asp:TextBox ID="DriverBoundsMonth" runat="server" CssClass="InputMargin" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red TableColAlign_Right">賠償費</td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbCompensateLoss11" runat="server" Text="桃園站" />
                    <br />
                    <asp:TextBox ID="eCompensateLoss11" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbCompensateLoss13" runat="server" Text="大溪站" />
                    <br />
                    <asp:TextBox ID="eCompensateLoss13" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbCompensateLoss15" runat="server" Text="大園站" />
                    <br />
                    <asp:TextBox ID="eCompensateLoss15" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbCompensateLoss16" runat="server" Text="觀音站" />
                    <br />
                    <asp:TextBox ID="eCompensateLoss16" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbCompensateLoss17" runat="server" Text="龍潭站" />
                    <br />
                    <asp:TextBox ID="eCompensateLoss17" runat="server" CssClass="InputMargin" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red TableColAlign_Right" />
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbCompensateLoss18" runat="server" Text="竹圍站" />
                    <br />
                    <asp:TextBox ID="eCompensateLoss18" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbCompensateLoss19" runat="server" Text="三峽站" />
                    <br />
                    <asp:TextBox ID="eCompensateLoss19" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbCompensateLoss20" runat="server" Text="新屋站" />
                    <br />
                    <asp:TextBox ID="eCompensateLoss20" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbCompensateLoss22" runat="server" Text="桃園公車站" />
                    <br />
                    <asp:TextBox ID="eCompensateLoss22" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbCompensateLoss23" runat="server" Text="中壢公車站" />
                    <br />
                    <asp:TextBox ID="eCompensateLoss23" runat="server" CssClass="InputMargin" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red TableColAlign_Right" />
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbCompensateLoss25" runat="server" Text="中壢站公車" />
                    <br />
                    <asp:TextBox ID="eCompensateLoss25" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red TableColAlign_Right" />
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red" />
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red TableColAlign_Right" />
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red" />
            </tr>
            <tr>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red TableColAlign_Right">專車收入</td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbTranIncome11" runat="server" Text="桃園站" />
                    <br />
                    <asp:TextBox ID="eTranIncome11" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbTranIncome13" runat="server" Text="大溪站" />
                    <br />
                    <asp:TextBox ID="eTranIncome13" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbTranIncome15" runat="server" Text="大園站" />
                    <br />
                    <asp:TextBox ID="eTranIncome15" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbTranIncome16" runat="server" Text="觀音站" />
                    <br />
                    <asp:TextBox ID="eTranIncome16" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbTranIncome17" runat="server" Text="龍潭站" />
                    <br />
                    <asp:TextBox ID="eTranIncome17" runat="server" CssClass="InputMargin" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red TableColAlign_Right" />
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbTranIncome18" runat="server" Text="竹圍站" />
                    <br />
                    <asp:TextBox ID="eTranIncome18" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbTranIncome19" runat="server" Text="三峽站" />
                    <br />
                    <asp:TextBox ID="eTranIncome19" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbTranIncome20" runat="server" Text="新屋站" />
                    <br />
                    <asp:TextBox ID="eTranIncome20" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbTranIncome22" runat="server" Text="桃園公車站" />
                    <br />
                    <asp:TextBox ID="eTranIncome22" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbTranIncome23" runat="server" Text="中壢公車站" />
                    <br />
                    <asp:TextBox ID="eTranIncome23" runat="server" CssClass="InputMargin" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red TableColAlign_Right" />
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:Label ID="lbTranIncome25" runat="server" Text="中壢站公車" />
                    <br />
                    <asp:TextBox ID="eTranIncome25" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red TableColAlign_Right" />
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red" />
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red TableColAlign_Right" />
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red" />
            </tr>
            <tr>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red TableColAlign_Right">每人服裝費</td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red">
                    <asp:TextBox ID="ClothCost" runat="server" CssClass="InputMargin" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red TableColAlign_Right" />
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red" />
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red TableColAlign_Right" />
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Red" />
            </tr>
            <tr>
                <td class="ColWidth-6Col ColHeight ColBorder text-Left-Blue TableColAlign_Right">
                    <asp:Button ID="bbSearch" Text="查詢" CssClass="button-Blue" runat="server" OnClick="bbSearch_Click" Width="40%" />
                    <asp:Button ID="bbClose" Text="結束" CssClass="button-Red" runat="server" OnClick="bbClose_Click" Width="40%" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plPrint" runat="server" GroupingText="預覽報表" CssClass="PrintPanel" Visible="false">
        <asp:Button ID="bbClosePrint" runat="server" CssClass="button-Red" Text="結束預覽" OnClick="bbClosePrint_Click" Width="100px" />
        <rsweb:ReportViewer ID="rvPrint" runat="server" BackColor="" ClientIDMode="AutoID" HighlightBackgroundColor=""
            InternalBorderColor="204, 204, 204" InternalBorderStyle="Solid" InternalBorderWidth="1px"
            LinkActiveColor="" LinkActiveHoverColor="" LinkDisabledColor=""
            PrimaryButtonBackgroundColor="" PrimaryButtonForegroundColor="" PrimaryButtonHoverBackgroundColor="" PrimaryButtonHoverForegroundColor=""
            SecondaryButtonBackgroundColor="" SecondaryButtonForegroundColor="" SecondaryButtonHoverBackgroundColor="" SecondaryButtonHoverForegroundColor=""
            SplitterBackColor="" ToolbarDividerColor="" ToolbarForegroundColor="" ToolbarForegroundDisabledColor="" ToolbarHoverBackgroundColor="" ToolbarHoverForegroundColor=""
            ToolBarItemBorderColor="" ToolBarItemBorderStyle="Solid" ToolBarItemBorderWidth="1px" ToolBarItemHoverBackColor="" ToolBarItemPressedBorderColor="51, 102, 153"
            ToolBarItemPressedBorderStyle="Solid" ToolBarItemPressedBorderWidth="1px" ToolBarItemPressedHoverBackColor="153, 187, 226" Width="100%">
            <LocalReport ReportPath="Report\TourCarIncomeP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
