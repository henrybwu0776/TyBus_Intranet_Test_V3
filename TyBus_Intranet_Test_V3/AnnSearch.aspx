<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="AnnSearch.aspx.cs" Inherits="TyBus_Intranet_Test_V3.AnnSearch" %>

<asp:Content ID="AnnSearchForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">公告資料查詢</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="發文單位：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Search_TextChanged" Width="30%" />
                    <asp:Label ID="eDepName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbEmpNo_Searech" runat="server" CssClass="text-Right-Blue" Text="發文人：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eEmpNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_Search_TextChanged" Width="30%" />
                    <asp:Label ID="eEmpName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbPostDate_Search" runat="server" CssClass="text-Right-Blue" Text="建立日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="ePostDate_S_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="titleText-S-Black" Text="～" Width="10%" />
                    <asp:TextBox ID="ePostDate_E_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbOpenDate_Search" runat="server" CssClass="text-Right-Blue" Text="有效日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eOpenDate_S_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="titleText-S-Black" Text="～" Width="10%" />
                    <asp:TextBox ID="eOpenDate_E_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lnPostTitle_Search" runat="server" CssClass="text-Right-Blue" Text="主旨:" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="ePostTitle_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbRemark_Search" runat="server" CssClass="text-Right-Blue" Text="本文：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:TextBox ID="eRemark_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Blue" OnClick="bbSearch_Click" Text="查詢" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" OnClick="bbClose_Click" Text="結束" Width="90%" />
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
        <asp:GridView ID="gridAnnList" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="AnnNo" DataSourceID="sdsAnnList_List"
            ForeColor="#333333" GridLines="None" PageSize="5" Width="100%">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="AnnNo" HeaderText="序號" ReadOnly="True" SortExpression="AnnNo" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="發文單位" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="EmpNo" HeaderText="EmpNo" SortExpression="EmpNo" Visible="False" />
                <asp:BoundField DataField="EmpName" HeaderText="發文人員" ReadOnly="True" SortExpression="EmpName" />
                <asp:BoundField DataField="PostDate" DataFormatString="{0:D}" HeaderText="發文日期" SortExpression="PostDate" />
                <asp:BoundField DataField="StartDate" DataFormatString="{0:D}" HeaderText="有效日(起)" SortExpression="StartDate" />
                <asp:BoundField DataField="EndDate" DataFormatString="{0:D}" HeaderText="有效日(迄)" SortExpression="EndDate" />
                <asp:BoundField DataField="PostTitle" HeaderText="主旨" SortExpression="PostTitle" />
            </Columns>
            <EditRowStyle BackColor="#999999" />
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#E9E7E2" />
            <SortedAscendingHeaderStyle BackColor="#506C8C" />
            <SortedDescendingCellStyle BackColor="#FFFDF8" />
            <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
        </asp:GridView>
        <asp:FormView ID="fvAnnList_Detail" runat="server" Width="100%" DataKeyNames="AnnNo" DataSourceID="sdsAnnList_Detail" OnDataBound="fvAnnList_Detail_DataBound">
            <ItemTemplate>
                <asp:Button ID="bbDownLoad_List" runat="server" CssClass="button-Blue" CausesValidation="false" OnClick="bbDownLoad_List_Click" Text="下載附檔" Width="120px" />
                <br />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAnnNo_List" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAnnNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AnnNo") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="發文單位：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="30%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbEmpNo_List" runat="server" CssClass="text-Right-Blue" Text="發文人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eEmpNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="30%" />
                            <asp:Label ID="eEmpName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPostTitle_List" runat="server" CssClass="text-Right-Blue" Text="發文主旨：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="ePostTitle_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PostTitle") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPostDate_List" runat="server" CssClass="text-Right-Blue" Text="發文日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="ePostDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PostDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStartDate_List" runat="server" CssClass="text-Right-Blue" Text="有效日(起)：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStartDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StartDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbEndDate_List" runat="server" CssClass="text-Right-Blue" Text="有效日(迄)：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eEndDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EndDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_Low ColBorder ColWidth-8Col" rowspan="2">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="公告本文：" Width="90%" />
                        </td>
                        <td class="MultiLine_Low ColBorder ColWidth-8Col" rowspan="2" colspan="5">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Height="98%" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:CheckBox ID="cbPostOpen_List" runat="server" CssClass="text-Left-Black" Text="發布在首頁" Width="90%" />
                            <asp:Label ID="ePostOpen_List" runat="server" Text='<%#Eval("PostOpen") %>' Visible="false" />
                        </td>
                    </tr>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbPostFiles_List" runat="server" CssClass="text-Right-Blue" Text="附檔檔名：" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="ePostFiles_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PostFiles") %>' Width="90%" />
                    </td>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' />
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
            </ItemTemplate>
        </asp:FormView>
        <asp:SqlDataSource ID="sdsAnnList_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT AnnNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, EmpNo, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.EmpNo)) AS EmpName, PostDate, StartDate, EndDate, PostTitle FROM AnnList AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
        <asp:SqlDataSource ID="sdsAnnList_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
            SelectCommand="SELECT AnnNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, EmpNo, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.EmpNo)) AS EmpName, PostDate, StartDate, EndDate, PostTitle, Remark, PostFiles, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = a.BuMan)) AS BuManName, BuDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName, ModifyDate, PostOpen FROM AnnList AS a WHERE (AnnNo = @AnnNo)">
            <SelectParameters>
                <asp:ControlParameter ControlID="gridAnnList" Name="AnnNo" PropertyName="SelectedValue" />
            </SelectParameters>
        </asp:SqlDataSource>
    </asp:Panel>
</asp:Content>
