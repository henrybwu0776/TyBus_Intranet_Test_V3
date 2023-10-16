<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DriverRegularTraining.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DriverRegularTraining" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="DriverRegularTrainingForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">駕駛員駕照及定期訓練證管理</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Start_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Start_Search_TextChanged" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text=" ～ " Width="5%" />
                    <asp:TextBox ID="eDepNo_End_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_End_Search_TextChanged" Width="40%" />
                    <br />
                    <asp:Label ID="eDepName_Start_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="eDepName_End_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbEmpNo_Search" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eEmpNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_Search_TextChanged" Width="45%" />
                    <asp:Label ID="eEmpName_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                </td>
                <td class="ColHeight ColWidth-8Col" colspan="2">
                    <asp:FileUpload ID="fuImportData" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Button ID="bbImportData" runat="server" CssClass="button-Black" Text="匯入" OnClick="bbImportData_Click" Width="45%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="4">
                    <asp:RadioButtonList ID="rbSearchMode" runat="server" CssClass="text-Left-Black" CellPadding="0" Width="90%" RepeatDirection="Horizontal">
                        <asp:ListItem Value="0" Text="所有駕駛員" Selected="True" />
                        <asp:ListItem Value="1" Text="過期尚未換發 (複訓) 人員" />
                        <asp:ListItem Value="2" Text="30天內需換發 (複訓) 人員" />
                    </asp:RadioButtonList>
                </td>                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="搜尋" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Red" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Black" Text="報表列印" OnClick="bbPrint_Click" Width="90%" />
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
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridDataList" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataKeyNames="EMPNO" DataSourceID="sdsDataList" GridLines="None" Width="100%">
            <Columns>
                <asp:BoundField DataField="DEPNO" HeaderText="DEPNO" SortExpression="DEPNO" Visible="false" />
                <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="EMPNO" HeaderText="員工編號" ReadOnly="True" SortExpression="EMPNO" />
                <asp:BoundField DataField="NAME" HeaderText="姓名" SortExpression="NAME" />
                <asp:BoundField DataField="LicenceType" HeaderText="LicenceType" SortExpression="LicenceType" Visible="false" />
                <asp:BoundField DataField="LicenceType_C" HeaderText="駕照種類" ReadOnly="True" SortExpression="LicenceType_C" />
                <asp:BoundField DataField="LicencingDay" DataFormatString="{0:d}" HeaderText="駕照發照日" SortExpression="LicencingDay" />
                <asp:BoundField DataField="LicenceCheck" DataFormatString="{0:d}" HeaderText="駕照審驗日" SortExpression="LicenceCheck" />
                <asp:BoundField DataField="DriverLicenceLimitDate" DataFormatString="{0:d}" HeaderText="駕照審驗期限" ReadOnly="True" SortExpression="DriverLicenceLimitDate" />
                <asp:BoundField DataField="TRATOOL" HeaderText="TRATOOL" SortExpression="TRATOOL" Visible="false" />
                <asp:BoundField DataField="TraTool_C" HeaderText="定期訓練種類" ReadOnly="True" SortExpression="TraTool_C" />
                <asp:BoundField DataField="TrainingGetDate" DataFormatString="{0:d}" HeaderText="定期訓練發證日" SortExpression="TrainingGetDate" />
                <asp:BoundField DataField="TrainingStopDate" DataFormatString="{0:d}" HeaderText="定期訓練到期日" SortExpression="TrainingStopDate" />
                <asp:BoundField DataField="TrainingLimitDate" DataFormatString="{0:d}" HeaderText="定期訓練審驗期限" ReadOnly="True" SortExpression="TrainingLimitDate" />

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
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
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
            <LocalReport ReportPath="Report\DriverRegularTrainingP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsDataList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DEPNO)) AS DepName, EMPNO, NAME, LicenceType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '人事資料檔      Employee        LicenceType') AND (CLASSNO = e.LicenceType)) AS LicenceType_C, LicencingDay, LicenceCheck, DATEADD(day, - 15, LicenceCheck) AS DriverLicenceLimitDate, TRATOOL, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (FKEY = '人事資料檔      EMPLOYEE        TRATOOL') AND (CLASSNO = e.TRATOOL)) AS TraTool_C, AUTONO AS TrainingGetDate, BBCALL AS TrainingStopDate, DATEADD(day, - 15, BBCALL) AS TrainingLimitDate FROM EMPLOYEE AS e WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
