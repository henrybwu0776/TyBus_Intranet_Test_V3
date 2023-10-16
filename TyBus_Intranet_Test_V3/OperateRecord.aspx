<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="OperateRecord.aspx.cs" Inherits="TyBus_Intranet_Test_V3.OperateRecord" %>

<asp:Content ID="OperateRecordForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">系統操作記錄</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbLoginID_Search" runat="server" CssClass="text-Right-Blue" Text="登入人員" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eLoginID_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eLoginID_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eLoginName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbOperateMode_Search" runat="server" CssClass="text-Right-Blue" Text="操作行為" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlOperateMode_Saerch" runat="server" CssClass="text-Left-Black" AutoPostBack="True" OnSelectedIndexChanged="ddlOperateMode_Saerch_SelectedIndexChanged" Width="95%">
                        <asp:ListItem Value="00" Text="" Selected="True" />
                        <asp:ListItem Value="01" Text="查詢資料" />
                        <asp:ListItem Value="02" Text="匯出檔案" />
                        <asp:ListItem Value="03" Text="列印報表" />
                        <asp:ListItem Value="04" Text="其他" />
                    </asp:DropDownList>
                    <asp:Label ID="eOperateMode_Search" runat="server" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbOperateDate_Search" runat="server" CssClass="text-Right-Blue" Text="操作日期" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eOperateDateS_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Blue" Text="～" />
                    <asp:TextBox ID="eOperateDateE_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbOperateFunction_Search" runat="server" CssClass="text-Right-Blue" Text="操作功能" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:DropDownList ID="ddlOperateFunctionMain_Search" runat="server" CssClass="text-Left-Blue" AutoPostBack="True" OnSelectedIndexChanged="ddlOperateFunctionMain_Search_SelectedIndexChanged" Width="50%" DataSourceID="dsOperateFunctionMain_Search" DataTextField="ClassTxt" DataValueField="ClassNo" />
                    <asp:DropDownList ID="ddlOperateFunctionSub_Search" runat="server" CssClass="text-Left-Blue" Width="40%" DataSourceID="dsOperateFunctionSub_Search" DataTextField="ControlCName" DataValueField="TargetPage" />
                </td>
                <td class="ColWidth-8Col" />
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" OnClick="bbSearch_Click" Text="查詢資料" Width="120px" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" OnClick="bbExcel_Click" Text="匯出 EXCEL" Width="120px" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" OnClick="bbClose_Click" Text="離開" Width="120px" />
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
    <asp:Panel ID="plShowData_List" runat="server" CssClass="ShowPanel-Detail_B">
        <asp:GridView ID="gridOperateRecord_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" CellPadding="3" DataKeyNames="IndexNo" DataSourceID="dsOperateRecord_List" GridLines="Horizontal" Width="100%" OnPageIndexChanging="gridOperateRecord_List_PageIndexChanging">
            <AlternatingRowStyle BackColor="#F7F7F7" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="IndexNo" HeaderText="IndexNo" InsertVisible="False" ReadOnly="True" SortExpression="IndexNo" Visible="False" />
                <asp:BoundField DataField="LoginID" HeaderText="登入帳號" SortExpression="LoginID" />
                <asp:BoundField DataField="EmpName" HeaderText="姓名" ReadOnly="True" SortExpression="EmpName" />
                <asp:BoundField DataField="ComputerName" HeaderText="電腦名稱 (IP)" SortExpression="ComputerName" />
                <asp:BoundField DataField="OperateDate" HeaderText="操作日期時間" SortExpression="OperateDate" />
                <asp:BoundField DataField="OperationNote" HeaderText="操作行為" SortExpression="OperationNote" />
            </Columns>
            <FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
            <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#F7F7F7" />
            <PagerStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" HorizontalAlign="Right" />
            <RowStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="#F7F7F7" />
            <SortedAscendingCellStyle BackColor="#F4F4FD" />
            <SortedAscendingHeaderStyle BackColor="#5A4C9D" />
            <SortedDescendingCellStyle BackColor="#D8D8F0" />
            <SortedDescendingHeaderStyle BackColor="#3E3277" />
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="plShowData_Detail" runat="server" CssClass="ShowPanel-Detail_C">
        <asp:FormView ID="fvOperateRecord_Detail" runat="server" DataSourceID="dsOperateRecord_Detail" Width="100%">            
            <ItemTemplate>
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColWidth-4Col">
                            <asp:Label ID="lbLoginID_List" runat="server" CssClass="text-Right-Blue" Text="登入帳號" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-4Col" colspan="3">
                            <asp:Label ID="eLoginID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LoginID") %>' Width="100%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-4Col">
                            <asp:Label ID="lbEmpName_List" runat="server" CssClass="text-Right-Blue" Text="姓名" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-4Col" colspan="3">
                            <asp:Label ID="eEmpName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="100%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-4Col">
                            <asp:Label ID="lbComputerName_List" runat="server" CssClass="text-Right-Blue" Text="電腦名稱 (IP)" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-4Col" colspan="3">
                            <asp:Label ID="eComputerName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ComputerName") %>' Width="100%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-4Col">
                            <asp:Label ID="lbOperateDate_List" runat="server" CssClass="text-Right-Blue" Text="操作日期時間" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-4Col" colspan="3">
                            <asp:Label ID="eOperateDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OperateDate") %>' Width="100%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-4Col">
                            <asp:Label ID="lbOperationNote_List" runat="server" CssClass="text-Right-Blue" Text="操作行為" Width="100%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-4Col" colspan="3" rowspan="6">
                            <asp:TextBox ID="eOperationNote_List" runat="server" TextMode="MultiLine" CssClass="text-Left-Black" Text='<%# Eval("OperationNote") %>' Width="100%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-4Col ColHeight" />
                    </tr>
                    <tr>
                        <td class="ColWidth-4Col ColHeight" />
                    </tr>
                    <tr>
                        <td class="ColWidth-4Col ColHeight" />
                    </tr>
                    <tr>
                        <td class="ColWidth-4Col ColHeight" />
                    </tr>
                    <tr>
                        <td class="ColWidth-4Col ColHeight" />
                    </tr>
                    <tr>
                        <td class="ColWidth-4Col" />
                        <td class="ColWidth-4Col" />
                        <td class="ColWidth-4Col" />
                        <td class="ColWidth-4Col" />
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="dsOperateFunctionMain_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT CAST(NULL AS varchar) AS ClassNo, CAST(NULL AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '網頁功能權限群組WebPermission   GroupID') ORDER BY ClassNo" />
    <asp:SqlDataSource ID="dsOperateFunctionSub_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT CAST(NULL AS varchar) AS GroupID, CAST(NULL AS varchar) AS TargetPage, CAST(NULL AS varchar) AS ControlCName UNION ALL SELECT GroupID, TargetPage, ControlCName FROM WebPermissionA ORDER BY GroupID" />
    <asp:SqlDataSource ID="dsOperateRecord_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT IndexNo, LoginID, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = OperateRecord.LoginID)) AS EmpName, ComputerName, OperateDate, OperationNote FROM OperateRecord WHERE (ISNULL(IndexNo, '') = '')"></asp:SqlDataSource>
    <asp:SqlDataSource ID="dsOperateRecord_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT LoginID, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = OperateRecord.LoginID)) AS EmpName, ComputerName, OperateDate, OperationNote FROM OperateRecord where IndexNo = @IndexNo">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridOperateRecord_List" Name="IndexNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
</asp:Content>
