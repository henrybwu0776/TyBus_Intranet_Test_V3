<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CarDeviceList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CarDeviceList" %>

<asp:Content ID="CarDeviceListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">車輛配備清冊匯出</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_S" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_E" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbTran_Type" runat="server" CssClass="text-Right-Blue" Text="狀態" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlTran_Type" runat="server" CssClass="text-Left-Black" Width="90%"
                        DataSourceID="sdsTran_Type" DataTextField="ClassTxt" DataValueField="ClassNo" OnSelectedIndexChanged="ddlTran_Type_SelectedIndexChanged" />
                    <br />
                    <asp:TextBox ID="eTran_Type" runat="server" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出 Excel" OnClick="bbExcel_Click" Width="90%" />
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
        <asp:GridView ID="gridCarDeviceList" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataKeyNames="Car_No" DataSourceID="sdsCarDeviceList" GridLines="None" PageSize="20">
            <Columns>
                <asp:BoundField DataField="Car_No" HeaderText="Car_No" ReadOnly="True" SortExpression="Car_No" Visible="False" />
                <asp:BoundField DataField="Car_ID" HeaderText="牌照號碼" SortExpression="Car_ID" />
                <asp:BoundField DataField="CompanyNo" HeaderText="CompanyNo" SortExpression="CompanyNo" Visible="False" />
                <asp:BoundField DataField="CompanyName" HeaderText="配屬車站" ReadOnly="True" SortExpression="CompanyName" />
                <asp:BoundField DataField="ProdDate" DataFormatString="{0:d}" HeaderText="出廠日期" SortExpression="ProdDate" />
                <asp:BoundField DataField="getlicdate" DataFormatString="{0:d}" HeaderText="領照日期" SortExpression="getlicdate" />
                <asp:BoundField DataField="Car_TypeID" HeaderText="型式" SortExpression="Car_TypeID" />
                <asp:BoundField DataField="Canon" HeaderText="污染標準" SortExpression="Canon" />
                <asp:BoundField DataField="Car_Class" HeaderText="Car_Class" SortExpression="Car_Class" Visible="False" />
                <asp:BoundField DataField="Car_Class_C" HeaderText="廠牌" ReadOnly="True" SortExpression="Car_Class_C" />
                <asp:BoundField DataField="CLASS" HeaderText="CLASS" SortExpression="CLASS" Visible="False" />
                <asp:BoundField DataField="Class_C" HeaderText="類別" ReadOnly="True" SortExpression="Class_C" />
                <asp:BoundField DataField="point" HeaderText="point" SortExpression="point" Visible="False" />
                <asp:BoundField DataField="Point_C" HeaderText="用途" ReadOnly="True" SortExpression="Point_C" />
                <asp:BoundField DataField="carlevel" HeaderText="carlevel" SortExpression="carlevel" Visible="False" />
                <asp:BoundField DataField="CarLevel_C" HeaderText="車輛等級" ReadOnly="True" SortExpression="CarLevel_C" />
                <asp:BoundField DataField="exceptional" HeaderText="exceptional" SortExpression="exceptional" Visible="False" />
                <asp:BoundField DataField="Exceptional_C" HeaderText="特殊車種" ReadOnly="True" SortExpression="Exceptional_C" />
                <asp:BoundField DataField="Tran_Type" HeaderText="Tran_Type" SortExpression="Tran_Type" Visible="False" />
                <asp:BoundField DataField="Tran_Type_C" HeaderText="狀態" ReadOnly="True" SortExpression="Tran_Type_C" />
                <asp:BoundField DataField="Tran_Date" DataFormatString="{0:d}" HeaderText="異動日期" SortExpression="Tran_Date" />
                <asp:BoundField DataField="sitqty" HeaderText="座位數" SortExpression="sitqty" />
                <asp:BoundField DataField="standqty" HeaderText="立位數" SortExpression="standqty" />
                <asp:BoundField DataField="bodyclass" HeaderText="車體廠牌" SortExpression="bodyclass" />
                <asp:BoundField DataField="horsepower" HeaderText="馬力" SortExpression="horsepower" />
                <asp:BoundField DataField="FREEZETYPE" HeaderText="冷氣廠牌" SortExpression="FREEZETYPE" />
                <asp:BoundField DataField="Fix3date" DataFormatString="{0:d}" HeaderText="下次三級日" SortExpression="Fix3date" />
                <asp:BoundField DataField="nextedate" DataFormatString="{0:d}" HeaderText="下次檢驗日" SortExpression="nextedate" />
                <asp:BoundField DataField="NCheckTerm" DataFormatString="{0:d}" HeaderText="檢驗期限" SortExpression="NCheckTerm" />
                <asp:BoundField DataField="stopDate" DataFormatString="{0:d}" HeaderText="停駛日" SortExpression="stopDate" />
                <asp:BoundField DataField="cancelDate" DataFormatString="{0:d}" HeaderText="繳銷日" SortExpression="cancelDate" />
                <asp:BoundField DataField="rddate" DataFormatString="{0:d}" HeaderText="補照日" SortExpression="rddate" />
                <asp:BoundField DataField="saleDate" DataFormatString="{0:d}" HeaderText="出售日" SortExpression="saleDate" />
                <asp:BoundField DataField="ReDay" DataFormatString="{0:d}" HeaderText="重新領牌日" SortExpression="ReDay" />
                <asp:BoundField DataField="remark" HeaderText="備註" SortExpression="remark" />
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
    <asp:SqlDataSource ID="sdsCarDeviceList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT Car_No, Car_ID, CompanyNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.CompanyNo)) AS CompanyName, ProdDate, getlicdate, Car_TypeID, Canon, Car_Class, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '車輛資料作業    Car_infoA       CAR_CLASS') AND (CLASSNO = a.Car_Class)) AS Car_Class_C, CLASS, (SELECT CLASSTXT FROM DBDICB AS DBDICB_5 WHERE (FKEY = '車輛資料作業    Car_infoA       CLASS') AND (CLASSNO = a.CLASS)) AS Class_C, point, (SELECT CLASSTXT FROM DBDICB AS DBDICB_4 WHERE (FKEY = '車輛資料作業    Car_infoA       POINT') AND (CLASSNO = a.point)) AS Point_C, carlevel, (SELECT CLASSTXT FROM DBDICB AS DBDICB_3 WHERE (FKEY = '車輛資料作業    Car_infoA       CARLEVEL') AND (CLASSNO = a.carlevel)) AS CarLevel_C, exceptional, (SELECT CLASSTXT FROM DBDICB AS DBDICB_2 WHERE (FKEY = '車輛資料作業    Car_infoA       EXCEPTIONAL') AND (CLASSNO = a.exceptional)) AS Exceptional_C, Tran_Type, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (FKEY = '車輛資料作業    Car_infoA       TRAN_TYPE') AND (CLASSNO = a.Tran_Type)) AS Tran_Type_C, Tran_Date, sitqty, standqty, bodyclass, horsepower, FREEZETYPE, Fix3date, nextedate, NCheckTerm, stopDate, cancelDate, rddate, saleDate, ReDay, remark FROM Car_infoA AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsTran_Type" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '車輛資料作業    Car_infoA       TRAN_TYPE') ORDER BY ClassNo"></asp:SqlDataSource>
</asp:Content>
