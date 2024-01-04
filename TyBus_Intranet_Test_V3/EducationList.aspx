<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EducationList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EducationList" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="EducationListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">職安教育訓練名冊</a>
    </div>
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4">
                    <asp:ListBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" Width="95%" Height="97%" SelectionMode="Multiple" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4">
                    <asp:Label ID="lbEmpType_Search" runat="server" CssClass="text-Right-Blue" Text="身分別" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" rowspan="4">
                    <asp:CheckBoxList ID="eEmpType_Search" runat="server" CssClass="text-Left-Black" Width="95%">
                        <asp:ListItem Text="內勤人員" Value="00" />
                        <asp:ListItem Text="修護人員" Value="10" />
                        <asp:ListItem Text="行車人員" Value="20" />
                    </asp:CheckBoxList>
                </td>
                <td class="colh ColBorder ColWidth-8Col">
                    <asp:Label ID="lbClassTitle_Search" runat="server" CssClass="text-Right-Blue" Text="課程名稱" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:TextBox ID="eClassTitle_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbTeacherName_Search" runat="server" CssClass="text-Right-Blue" Text="講師姓名" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eTeacherName_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder  ColWidth-8Col">
                    <asp:Label ID="lbLocation_Search" runat="server" CssClass="text-Right-Blue" Text="上課地點" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eLocation_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbClassDate_Search" runat="server" CssClass="text-Right-Blue" Text="上課日期" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eClassDate_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbClassTime_Search" runat="server" CssClass="text-Right-Blue" Text="上課時間" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eClassTime_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                 </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" />
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Button ID="bbPreview" runat="server" CssClass="button-Black" Text="預覽報表" OnClick="bbPreview_Click" Width="120px" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="120px" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" Text="離開" OnClick="bbExit_Click" Width="120px" />
                </td>                
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col" />
            </tr>
        </table>
    </asp:Panel>
        <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCLoseReport" runat="server" CssClass="button-Red" Text="關閉報表" OnClick="bbCLoseReport_Click" Width="120px" />
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
            <LocalReport ReportPath="Report\EducationListP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
