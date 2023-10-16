<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CarInfoExport.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CarInfoExport" %>

<asp:Content ID="CarInfoExportForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">車輛資料匯出作業</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eDepNo_S" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_E" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbPoint" runat="server" CssClass="text-Right-Blue" Text="用途：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:DropDownList ID="ddlPoint_S" runat="server" CssClass="text-Left-Black" Width="40%" AutoPostBack="True"
                        DataSourceID="sdsPoint_S" DataTextField="CLASSTXT" DataValueField="CLASSNO"
                        OnSelectedIndexChanged="ddlPoint_S_SelectedIndexChanged" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:DropDownList ID="ddlPoint_E" runat="server" CssClass="text-Left-Black" Width="40%" AutoPostBack="True"
                        DataSourceID="sdsPoint_E" DataTextField="CLASSTXT" DataValueField="CLASSNO"
                        OnSelectedIndexChanged="ddlPoint_E_SelectedIndexChanged" />
                    <br />
                    <asp:TextBox ID="ePoint_S" runat="server" Visible="false" />
                    <asp:TextBox ID="ePoint_E" runat="server" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="coolh ColBorder ColWidth-6Col">
                    <asp:Label ID="lbCarYearCal" runat="server" CssClass="text-Right-Blue" Text="車齡計算基準日：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eCarYearCal" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbCarState" runat="server" CssClass="text-Right-Blue" Text="狀態：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:DropDownList ID="eCarState" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbClass" runat="server" CssClass="text-Right-Blue" Text="車型：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:DropDownList ID="eClass" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbExceptional" runat="server" CssClass="text-Right-Blue" Text="特殊車種：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:DropDownList ID="eExceptional" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出 Excel" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
                <td class="ColWidth-6Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridCarInfo" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="Car_No" DataSourceID="sdsCarInfo" GridLines="None" PageSize="20" Width="100%">
            <Columns>
                <asp:BoundField DataField="CompanyNo" HeaderText="CompanyNo" SortExpression="CompanyNo" Visible="False" />
                <asp:BoundField DataField="CompanyName" HeaderText="站別" ReadOnly="True" SortExpression="CompanyName" />
                <asp:BoundField DataField="Car_ID" HeaderText="牌照號碼" SortExpression="Car_ID" />
                <asp:BoundField DataField="Car_No" HeaderText="車輛代號" ReadOnly="True" SortExpression="Car_No" />
                <asp:BoundField DataField="Car_Class" HeaderText="Car_Class" ReadOnly="True" SortExpression="Car_Class" Visible="False" />
                <asp:BoundField DataField="Brand_C" HeaderText="廠牌" ReadOnly="True" SortExpression="Brand_C" />
                <asp:BoundField DataField="Car_TypeID" HeaderText="型式" ReadOnly="True" SortExpression="Car_TypeID" />
                <asp:BoundField DataField="ProdDate" DataFormatString="{0:d}" HeaderText="ProdDate" SortExpression="ProdDate" Visible="False" />
                <asp:BoundField DataField="ProdDate_Y" HeaderText="出廠年份" ReadOnly="True" SortExpression="ProdDate_Y" />
                <asp:BoundField DataField="ProdDate_M" HeaderText="出廠月份" ReadOnly="True" SortExpression="ProdDate_M" />
                <asp:BoundField DataField="CarYM" HeaderText="車齡" SortExpression="sitqty" />
                <asp:BoundField DataField="sitqty" HeaderText="座位數" SortExpression="sitqty" />
                <asp:BoundField DataField="Tran_Type" HeaderText="Tran_Type" SortExpression="Tran_Type" Visible="False" />
                <asp:BoundField DataField="Tran_Type_C" HeaderText="狀態" ReadOnly="True" SortExpression="Tran_Type_C" />
                <asp:BoundField DataField="point" HeaderText="point" SortExpression="point" Visible="False" />
                <asp:BoundField DataField="Point_C" HeaderText="用途" ReadOnly="True" SortExpression="Point_C" />
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
    <asp:SqlDataSource ID="sdsCarInfo" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT CompanyNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.CompanyNo)) AS CompanyName, Car_ID, Car_No, ProdDate, YEAR(ProdDate) AS ProdDate_Y, MONTH(ProdDate) AS ProdDate_M, CAST(DATEDIFF(Month, ProdDate, GETDATE()) / 12 AS varchar) + '年' + CAST(DATEDIFF(Month, ProdDate, GETDATE()) % 12 AS varchar) + '月' AS CarYM, sitqty, Tran_Type, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.Tran_Type) AND (FKEY = '車輛資料作業    Car_infoA       TRAN_TYPE')) AS Tran_Type_C, point, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (CLASSNO = a.point) AND (FKEY = '車輛資料作業    Car_infoA       POINT')) AS Point_C FROM Car_infoA AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsPoint_S" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '車輛資料作業    Car_infoA       POINT') ORDER BY CLASSNO"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsPoint_E" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '車輛資料作業    Car_infoA       POINT') ORDER BY CLASSNO"></asp:SqlDataSource>
</asp:Content>
