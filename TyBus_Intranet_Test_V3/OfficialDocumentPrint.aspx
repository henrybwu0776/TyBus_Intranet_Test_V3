<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="OfficialDocumentPrint.aspx.cs" Inherits="TyBus_Intranet_Test_V3.OfficialDocumentPrint" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="OfficialDocumentPrintForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">公文統計列印</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" GroupingText="篩選條件" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDocDep_Search" runat="server" CssClass="text-Right-Blue" Text="收發單位：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDocDep_Search" runat="server" AutoPostBack="true" CssClass="text-Left-Black" OnTextChanged="eDocDep_Search_TextChanged" Width="45%" />
                    <asp:Label ID="eDocDepName_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDocNo_Search" runat="server" CssClass="text-Right-Blue" Text="收發文字號：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:TextBox ID="eDocYears_Search" runat="server" CssClass="text-Left-Black" Width="15%" />
                    <asp:Label ID="lbDocSplit1" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:DropDownList ID="ddlDocFirstWord_Search" runat="server" CssClass="text-Left-Black"
                        OnSelectedIndexChanged="ddlDocFirstWord_Search_SelectedIndexChanged" Width="15%">
                    </asp:DropDownList>
                    <asp:Label ID="lbDocSplit2" runat="server" CssClass="text-Left-Black" Text=" 字第 " />
                    <asp:TextBox ID="eDocNo_Start_Search" runat="server" CssClass="text-Left-Black" Width="20%" />
                    <asp:Label ID="lbDocSplit3" runat="server" CssClass="text-Left-Black" Text=" ～ " />
                    <asp:TextBox ID="eDocNo_End_Search" runat="server" CssClass="text-Left-Black" Width="20%" />
                    <asp:Label ID="lbDocSplit4" runat="server" CssClass="text-Left-Black" Text="　號" />
                    <br />
                    <asp:Label ID="eDocFirstWord_Search" runat="server" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDocDate" runat="server" CssClass="text-Right-Blue" Text="收發日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDocDate_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbDocSplit5" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDocDate_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbStoreDate_Search" runat="server" CssClass="text-Right-Blue" Text="歸檔日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eStoreDate_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbDocSplit6" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eStoreDate_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
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
    <asp:Panel ID="plReport" runat="server" CssClass="PrintPanel">
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
            <LocalReport ReportPath="Report\OfficialDocumentP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridOfficialDocumentP" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="DocIndex"
            DataSourceID="sdsOfficialDocumentP" GridLines="None" Width="100%">
            <Columns>
                <asp:BoundField DataField="DocIndex" HeaderText="DocIndex" ReadOnly="True" SortExpression="DocIndex" Visible="False" />
                <asp:BoundField DataField="DocDep_C" HeaderText="單位" ReadOnly="True" SortExpression="DocDep_C" />
                <asp:BoundField DataField="BuildMan_C" HeaderText="建檔人" ReadOnly="True" SortExpression="BuildMan_C" />
                <asp:BoundField DataField="DocDate" DataFormatString="{0:d}" HeaderText="收發文日期" SortExpression="DocDate" />
                <asp:BoundField DataField="DocDep" HeaderText="DocDep" SortExpression="DocDep" Visible="False" />
                <asp:BoundField DataField="DocFirstWord" HeaderText="DocFirstWord" SortExpression="DocFirstWord" Visible="False" />
                <asp:BoundField DataField="DocFirstWord_C" HeaderText="DocFirstWord_C" ReadOnly="True" SortExpression="DocFirstWord_C" Visible="False" />
                <asp:BoundField DataField="DocNo" HeaderText="DocNo" SortExpression="DocNo" Visible="False" />
                <asp:BoundField DataField="DocSourceUnit" HeaderText="來文機關" SortExpression="DocSourceUnit" />
                <asp:BoundField DataField="DocType" HeaderText="DocType" SortExpression="DocType" Visible="False" />
                <asp:BoundField DataField="DocType_C" HeaderText="文別" ReadOnly="True" SortExpression="DocType_C" />
                <asp:BoundField DataField="DocTitle" HeaderText="事由" SortExpression="DocTitle" />
                <asp:BoundField DataField="Undertaker_C" HeaderText="承辦人" ReadOnly="True" SortExpression="Undertaker_C" />
                <asp:BoundField DataField="Undertaker" HeaderText="Undertaker" SortExpression="Undertaker" Visible="False" />
                <asp:BoundField DataField="OutsideDocFirstWord" HeaderText="OutsideDocFirstWord" SortExpression="OutsideDocFirstWord" Visible="False" />
                <asp:BoundField DataField="OutsideDocNo" HeaderText="OutsideDocNo" SortExpression="OutsideDocNo" Visible="False" />
                <asp:BoundField DataField="Attachement" HeaderText="Attachement" SortExpression="Attachement" Visible="False" />
                <asp:BoundField DataField="Implementation" HeaderText="Implementation" SortExpression="Implementation" Visible="False" />
                <asp:BoundField DataField="BuildMan" HeaderText="BuildMan" SortExpression="BuildMan" Visible="False" />
                <asp:BoundField DataField="BuildDate" HeaderText="BuildDate" SortExpression="BuildDate" Visible="False" />
                <asp:BoundField DataField="Remark" HeaderText="Remark" SortExpression="Remark" Visible="False" />
                <asp:BoundField DataField="StoreDate" DataFormatString="{0:d}" HeaderText="歸檔日期" SortExpression="StoreDate" />
                <asp:BoundField DataField="StoreMan" HeaderText="StoreMan" SortExpression="StoreMan" Visible="False" />
                <asp:BoundField DataField="StoreMan_C" HeaderText="歸檔人" ReadOnly="True" SortExpression="StoreMan_C" />
                <asp:BoundField DataField="Remark_Store" HeaderText="Remark_Store" SortExpression="Remark_Store" Visible="False" />
            </Columns>
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
    <asp:SqlDataSource ID="sdsOfficialDocumentP" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT DocIndex, DocDate, DocYears, DocDep, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = OfficialDocument.DocDep)) AS DocDep_C, DocFirstWord, (SELECT DocFirstCWord FROM DOCFirstWord WHERE (FWNo = OfficialDocument.DocFirstWord)) AS DocFirstWord_C, DocNo, DocSourceUnit, DocType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = CAST('公文收發登記' AS char(16)) + CAST('OfficialDocument' AS char(16)) + 'DocType') AND (CLASSNO = OfficialDocument.DocType)) AS DocType_C, DocTitle, Undertaker, (SELECT NAME FROM EMPLOYEE WHERE (LEAVEDAY IS NULL) AND (EMPNO = OfficialDocument.Undertaker)) AS Undertaker_C, OutsideDocFirstWord, OutsideDocNo, Attachement, Implementation, BuildMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (LEAVEDAY IS NULL) AND (EMPNO = OfficialDocument.BuildMan)) AS BuildMan_C, BuildDate, Remark, StoreDate, StoreMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (LEAVEDAY IS NULL) AND (EMPNO = OfficialDocument.StoreMan)) AS StoreMan_C, Remark_Store FROM OfficialDocument WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
