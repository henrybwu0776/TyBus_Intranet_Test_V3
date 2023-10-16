<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CreateNameList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CreateNameList" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="CreateNameListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">福利品發放名冊</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbListType" runat="server" CssClass="text-Right-Blue" Text="發放種類：" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="7">
                    <asp:RadioButtonList ID="rbListType" runat="server" CssClass="text-Left-Black" RepeatColumns="7" Width="100%"
                        OnSelectedIndexChanged="rbListType_SelectedIndexChanged" AutoPostBack="True">
                        <asp:ListItem Value="0" Text="春節福利品" Selected="True" />
                        <asp:ListItem Value="1" Text="端午福利品" />
                        <asp:ListItem Value="2" Text="中秋福利品" />
                        <asp:ListItem Value="3" Text="紅包" />
                        <asp:ListItem Value="4" Text="尾牙" />
                        <asp:ListItem Value="5" Text="秋節月餅" />
                        <asp:ListItem Value="6" Text="制服領用" />
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDateRange" runat="server" CssClass="text-Right-Blue" Text="半份起迄日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:TextBox ID="eStartDate_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eEndDate_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbGiftYear" runat="server" CssClass="text-Right-Blue" Text="年度：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eGiftYear" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbMoneyPay" runat="server" CssClass="text-Right-Blue" Text="發放金額：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eMoneyPay" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="MultiLine-High ColBorder ColWidth-8Col">
                    <asp:Label ID="lbNote" runat="server" CssClass="text-Right-Blue" Text="發放說明：" Width="90%" />
                </td>
                <td class="MultiLine-High ColBorder ColWidth-8Col" colspan="7">
                    <asp:TextBox ID="eNote" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Height="95%" Width="97%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                    <asp:Label ID="lbRemark_Search" runat="server" CssClass="errorMessageText" Text="發放對象：報表產生當下在職人員 (不含留停、離職、入伍)" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbPreview" runat="server" CssClass="button-Black" Text="預覽" OnClick="bbPreview_Click" Width="90%" />
                </td>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbExportExcel" runat="server" CssClass="button-Black" Text="匯出 EXCEL" OnClick="bbExportExcel_Click" Width="90%" />
                </td>
                <td class="ColWidth-8Col">
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
        <asp:GridView ID="gridShowData" runat="server" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" GridLines="None">
            <FooterStyle BackColor="#C6C3C6" ForeColor="Black" />
            <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#E7E7FF" />
            <PagerStyle BackColor="#C6C3C6" ForeColor="Black" HorizontalAlign="Right" />
            <RowStyle BackColor="#DEDFDE" ForeColor="Black" />
            <SelectedRowStyle BackColor="#9471DE" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F1F1F1" />
            <SortedAscendingHeaderStyle BackColor="#594B9C" />
            <SortedDescendingCellStyle BackColor="#CAC9C9" />
            <SortedDescendingHeaderStyle BackColor="#33276A" />
        </asp:GridView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsGiftNameList_Gift" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DEPNO)) AS DepName, EMPNO, NAME, TITLE, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = e.TITLE) AND (FKEY = '員工資料        EMPLOYEE        TITLE')) AS Title_C, ASSUMEDAY, CAST('' AS nvarchar) AS GiftType, CAST('說明' AS nvarchar) AS GiftNote, CAST('' AS nvarchar) AS GiftTitle, CAST(0 AS integer) AS FullGift, CAST(0 AS integer) AS HalfGift FROM EMPLOYEE AS e WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsGiftNameList_Money" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DEPNO)) AS DepName, EMPNO, NAME, TITLE, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = e.TITLE) AND (FKEY = '員工資料        EMPLOYEE        TITLE')) AS Title_C, ASSUMEDAY, CAST('' AS nvarchar) AS GiftNote, CAST(500 AS float) AS GiftPay, CAST('108 年 紅包發放名冊' AS nvarchar) AS GiftTitle, CAST('' AS nvarchar) AS StampPlace FROM (SELECT DEPNO, EMPNO, NAME, TITLE, ASSUMEDAY FROM EMPLOYEE WHERE (1 &lt;&gt; 1) UNION ALL SELECT DEPNO, EMPNO, NAME, TITLE, ASSUMEDAY FROM EMPLOYEE AS EMPLOYEE_1 WHERE (1 &lt;&gt; 1)) AS e WHERE (EMPNO &lt;&gt; 'supervisor') ORDER BY TITLE, DEPNO, EMPNO"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsGiftNameList_Uniform" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DEPNO)) AS DepName, EMPNO, NAME, TITLE, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = e.TITLE) AND (FKEY = '員工資料        EMPLOYEE        TITLE')) AS Title_C, ASSUMEDAY, CAST('' AS nvarchar) AS GiftNote, CAST('108 年 制服領用發放名冊' AS nvarchar) AS GiftTitle, CAST('' AS nvarchar) AS StampPlace FROM (SELECT DEPNO, EMPNO, NAME, TITLE, ASSUMEDAY FROM EMPLOYEE WHERE (1 &lt;&gt; 1) UNION ALL SELECT DEPNO, EMPNO, NAME, TITLE, ASSUMEDAY FROM EMPLOYEE AS EMPLOYEE_1 WHERE (1 &lt;&gt; 1)) AS e WHERE (EMPNO &lt;&gt; 'supervisor') ORDER BY TITLE, DEPNO, EMPNO"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsGiftNameList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"></asp:SqlDataSource>

    <asp:Panel ID="plReport" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbClosePreview" runat="server" CssClass="button-Red" Text="結束預覽" OnClick="bbClosePreview_Click" Width="120px" />
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
            ToolBarItemPressedHoverBackColor="153, 187, 226" Height="600px" Width="100%" PageCountMode="Actual" ShowBackButton="False" ShowFindControls="False" ShowParameterPrompts="False" ShowRefreshButton="False" ZoomMode="PageWidth">
            <LocalReport ReportPath="Report\GiftNameList.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
