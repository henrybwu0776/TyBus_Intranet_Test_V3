<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="TicketTrans.aspx.cs" Inherits="TyBus_Intranet_Test_V3.TicketTrans" %>

<asp:Content ID="TicketTransForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">電子票證營收轉入</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbBuDate_Search" runat="server" CssClass="text-Right-Blue" Text="清分日期" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eBuDate_S_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eBuDate_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbExcelFile_Search" runat="server" CssClass="text-Right-Blue" Text="轉入檔案" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:FileUpload ID="fuTicketTrans" runat="server" CssClass="text-Left-Blue" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" colspan="3" />
                <td class="ColHeight ColWidth-6Col" colspan="3">                    
                    <asp:Button ID="bbOK_Search" runat="server" CssClass="button-Black" OnClick="bbOK_Search_Click" Text="查詢清分資料" Width="20%" />
                    <asp:Button ID="bbImport_Search" runat="server" CssClass="button-Blue" OnClick="bbImport_Search_Click" Text="匯入清分資料" Width="20%" />
                    <asp:Button ID="bbClear_Search" runat="server" CssClass="button-Black" OnClick="bbClear_Search_Click" Text="刪除清分資料" Width="20%" /> 
                    <asp:Button ID="bbClose_Search" runat="server" CssClass="button-Red" OnClick="bbClose_Search_Click" Text="離開" Width="20%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridCPSOutG" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px" CellPadding="4" DataKeyNames="IndexNo" DataSourceID="sdsCPSOutG_List" Width="100%">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="IndexNo" HeaderText="序號" ReadOnly="True" SortExpression="IndexNo" Visible="False" />
                <asp:BoundField DataField="budate" DataFormatString="{0:D}" HeaderText="清分日期" SortExpression="budate" />
                <asp:BoundField DataField="date" DataFormatString="{0:D}" HeaderText="營運日期" SortExpression="date" />
                <asp:BoundField DataField="depno" HeaderText="站別代碼" SortExpression="depno" />
                <asp:BoundField DataField="linesno" HeaderText="驗票機路線代碼" SortExpression="linesno" />
                <asp:BoundField DataField="car_id" HeaderText="牌照號碼" SortExpression="car_id" />
                <asp:BoundField DataField="car_no" HeaderText="car_no" SortExpression="car_no" Visible="False" />
                <asp:BoundField DataField="driver" HeaderText="駕駛員工號" SortExpression="driver" />
            </Columns>
            <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
            <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
            <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
            <RowStyle BackColor="White" ForeColor="#003399" />
            <SelectedRowStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
            <SortedAscendingCellStyle BackColor="#EDF6F6" />
            <SortedAscendingHeaderStyle BackColor="#0D4AC4" />
            <SortedDescendingCellStyle BackColor="#D6DFDF" />
            <SortedDescendingHeaderStyle BackColor="#002876" />
        </asp:GridView>
        <asp:FormView ID="fvCPSOutG" runat="server" DataSourceID="sdsCPSOutG_Detail" Width="100%">
            <ItemTemplate>
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eIndexNo_List" runat="server" Text='<%# Eval("IndexNo") %>' Visible="false" />
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="清分日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0: yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDEpNo_List" runat="server" CssClass="text-Right-Blue" Text="站別" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriver_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDriver_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver") %>' Width="35%" />
                            <asp:Label ID="eDriverName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDate_List" runat="server" CssClass="text-Right-Blue" Text="營運日期" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Date", "{0: yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCarID_List" runat="server" CssClass="text-Right-Blue" Text="牌照號碼" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCarID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLinesNo_List" runat="server" CssClass="text-Right-Blue" Text="驗票機路線代碼" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLinesNo_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("LinesNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbHCount_List" runat="server" CssClass="text-Right-Blue" Text="總載客人數" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eHCount_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("HCount") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbOAMT_List" runat="server" CssClass="text-Right-Blue" Text="營運金額 (普卡)" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eOAmt_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("OAmt") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbOPiece_List" runat="server" CssClass="text-Right-Blue" Text="營運點數 (聯營卡)" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eOPiece_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("OPiece") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPCount_List" runat="server" CssClass="text-Right-Blue" Text="普卡人數" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="ePCount_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("PCount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLCount_List" runat="server" CssClass="text-Right-Blue" Text="聯營卡人數" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLCount_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("LCount") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSPiece_List" runat="server" CssClass="text-Right-Blue" Text="學生卡點數" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSPiece_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("SPiece") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSCount_List" runat="server" CssClass="text-Right-Blue" Text="學生卡人數" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSCount_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("SCount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbOldPiece_List" runat="server" CssClass="text-Right-Blue" Text="敬老卡點數" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eOldPiece_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("OldPiece") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbOldCount_List" runat="server" CssClass="text-Right-Blue" Text="敬老卡人數" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eOldCount_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("OldCount") %>' Width="95%" />
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
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsCPSOutG_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" 
        SelectCommand="SELECT CONVERT (char(8), budate, 112) + CONVERT (char(8), date, 112) + CAST(depno AS char(12)) + CAST(linesno AS char(12)) + CAST(car_id AS char(12)) + CAST(car_no AS char(12)) + CAST(driver AS char(12)) AS IndexNo, budate, date, depno, linesno, car_id, car_no, driver FROM CPSOUTG WHERE (ISNULL(budate, '') = '')"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsCPSOutG_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" 
        SelectCommand="SELECT CONVERT (char(8), budate, 112) + CONVERT (char(8), date, 112) + CAST(depno AS char(12)) + CAST(linesno AS char(12)) + CAST(car_id AS char(12)) + CAST(car_no AS char(12)) + CAST(driver AS char(12)) AS IndexNo, 
       budate, date, depno, (select [Name] from Department where DepNo = CPSOutG.DepNo) DepName, linesno, car_id, car_no, 
	   driver, (select [Name] from Employee where EmpNo = CPSOutG.Driver) DriverName, Oamt, Opiece, Spiece, Hcount, Scount, Post_id, Post_flag, Postdate, oldPiece, PCount, LCount, oldCount 
  FROM CPSOUTG 
 WHERE (CONVERT (char(8), budate, 112) + CONVERT (char(8), date, 112) + CAST(depno AS char(12)) + CAST(linesno AS char(12)) + CAST(car_id AS char(12)) + CAST(car_no AS char(12)) + CAST(driver AS char(12)) = @IndexNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="GridCPSOutG" Name="IndexNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
</asp:Content>
