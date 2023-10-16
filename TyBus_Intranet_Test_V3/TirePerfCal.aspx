<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="TirePerfCal.aspx.cs" Inherits="TyBus_Intranet_Test_V3.TirePerfCal" %>

<asp:Content ID="TirePerfCalForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">輪胎公里數展開</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalYM" runat="server" CssClass="text-Right-Blue" Text="月結年月：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eCalYear_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text=" 年" Width="10%" />
                    <asp:TextBox ID="eCalMonth_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text=" 月" Width="10%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbTireNo_Search" runat="server" CssClass="text-Right-Blue" Text="輪胎號碼：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eTireNo_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="95%" />
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
        <asp:GridView ID="gridTirePerfCalMain" runat="server" BackColor="White"
            BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" GridLines="None"
            AutoGenerateColumns="False" DataKeyNames="TIRENO" DataSourceID="sdsTirePerfCalMain" AllowPaging="True" PageSize="5"
            OnSelectedIndexChanged="gridTirePerfCalMain_SelectedIndexChanged" Width="100%">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="TIRENO" HeaderText="輪胎號碼" ReadOnly="True" SortExpression="TIRENO" />
                <asp:BoundField DataField="DEPNO" HeaderText="部門編號" SortExpression="DEPNO" />
                <asp:BoundField DataField="depname" HeaderText="部門名稱" ReadOnly="True" SortExpression="depname" />
                <asp:BoundField DataField="DATES" DataFormatString="{0:D}" HeaderText="起始日" SortExpression="DATES" />
                <asp:BoundField DataField="DATEE" DataFormatString="{0:D}" HeaderText="拆卸日" SortExpression="DATEE" />
                <asp:BoundField DataField="TIREKM" HeaderText="拆卸公里" SortExpression="TIREKM" />
                <asp:BoundField DataField="brand_name" HeaderText="廠牌" ReadOnly="True" SortExpression="brand_name" />
                <asp:BoundField DataField="spec_name" HeaderText="規格" ReadOnly="True" SortExpression="spec_name" />
                <asp:BoundField DataField="type_name" HeaderText="輪胎類別" ReadOnly="True" SortExpression="type_name" />
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
    <asp:Panel ID="plResultData" runat="server" CssClass="ShowPanel-Detail">
        <asp:Label ID="lbExcelVersion" runat="server" CssClass="text-Right-Blue" Text="EXCEL 版本：" Width="200px" Visible="false" />
        <asp:DropDownList ID="ddlExcelVersion" runat="server" CssClass="text-Left-Black" Width="200px" Visible="false">
            <asp:ListItem Value="OldExcel" Text="Excel 97-2003" Selected="True" />
            <asp:ListItem Value="NewExcel" Text="Excel 2010" />
        </asp:DropDownList>
        <asp:Button ID="bbOutToExcel" runat="server" CssClass="button-Red" Text="匯出 EXCEL" OnClick="bbOutToExcel_Click" Width="120px" />
        <br />
        <asp:GridView ID="gridResultData" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#DEDFDE" BorderStyle="None" BorderWidth="1px"
            CellPadding="4" DataSourceID="sdsResultData" ForeColor="Black" GridLines="Vertical" PageSize="5" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:BoundField DataField="budate" DataFormatString="{0:d}" HeaderText="行車日期" SortExpression="budate" />
                <asp:BoundField DataField="Car_id" HeaderText="牌照號碼" SortExpression="Car_id" />
                <asp:BoundField DataField="KM_Sum" HeaderText="公里小計" ReadOnly="True" SortExpression="KM_Sum" />
                <asp:BoundField DataField="CBus1Km" HeaderText="班車" SortExpression="CBus1Km" />
                <asp:BoundField DataField="CBus2Km" HeaderText="公車" SortExpression="CBus2Km" />
                <asp:BoundField DataField="CRentTraKm" HeaderText="合約車" SortExpression="CRentTraKm" />
                <asp:BoundField DataField="CRentAKm" HeaderText="區間租車" SortExpression="CRentAKm" />
                <asp:BoundField DataField="CRentBKm" HeaderText="遊覽車" SortExpression="CRentBKm" />
                <asp:BoundField DataField="CBus3Km" HeaderText="專車" SortExpression="CBus3Km" />
                <asp:BoundField DataField="CBus4Km" HeaderText="聯營" SortExpression="CBus4Km" />
                <asp:BoundField DataField="CBus5Km" HeaderText="非營運" SortExpression="CBus5Km" />
            </Columns>
            <FooterStyle BackColor="#CCCC99" />
            <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
            <RowStyle BackColor="#F7F7DE" />
            <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#FBFBF2" />
            <SortedAscendingHeaderStyle BackColor="#848384" />
            <SortedDescendingCellStyle BackColor="#EAEAD3" />
            <SortedDescendingHeaderStyle BackColor="#575357" />
        </asp:GridView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsTirePerfCalMain" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT A.TIRENO, A.DEPNO, (SELECT NAME FROM Department WHERE (DEPNO = A.DEPNO)) AS depname, A.DATES, A.DATEE, A.TIREKM, (SELECT CLASSTXT FROM Dbdicb WHERE (FKEY = '輪胎資料        Tires           BRAND') AND (CLASSNO = B.Brand)) AS brand_name, (SELECT CLASSTXT FROM Dbdicb AS Dbdicb_2 WHERE (FKEY = '輪胎資料        Tires           SPEC') AND (CLASSNO = B.Spec)) AS spec_name, (SELECT CLASSTXT FROM Dbdicb AS Dbdicb_1 WHERE (FKEY = '輪胎資料        Tires           TYPE') AND (CLASSNO = B.Type)) AS type_name FROM TirePerf AS A INNER JOIN Tires AS B ON A.TIRENO = B.TireNo WHERE (1 = 1) ORDER BY A.TIRENO"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsResultData" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST(NULL AS DateTime) AS BuDate, CAST('' AS varchar) AS Car_ID, CAST(0 AS float) AS KM_Sum, CAST(0 AS float) AS CBus1KM, CAST(0 AS float) AS CBus2KM, CAST(0 AS float) AS CRentTraKM, CAST(0 AS float) AS CRentAKM, CAST(0 AS float) AS CRentBKM, CAST(0 AS float) AS CBus3KM, CAST(0 AS float) AS CBus4KM, CAST(0 AS float) AS CBus5KM FROM RsCarTotal WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
