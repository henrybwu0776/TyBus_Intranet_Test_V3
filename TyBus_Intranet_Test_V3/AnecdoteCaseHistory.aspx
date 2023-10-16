<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="AnecdoteCaseHistory.aspx.cs" Inherits="TyBus_Intranet_Test_V3.AnecdoteCaseHistory" %>

<asp:Content ID="AnecdoteCaseHistoryForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">肇事案件異動記錄</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbBuildDate_Search" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eBuildDate_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eBuildDate_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="4">
                    <asp:TextBox ID="eDepNo_Start_Search" runat="server" CssClass="text-Left-Black" Width="10%" AutoPostBack="True"
                        OnTextChanged="eDepNo_Start_Search_TextChanged" />
                    <asp:Label ID="eDepName_Start_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbSplit2" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_End_Search" runat="server" CssClass="text-Left-Black" Width="10%" AutoPostBack="True"
                        OnTextChanged="eDepNo_End_Search_TextChanged" />
                    <asp:Label ID="eDepName_End_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
                </td>
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
    <asp:Panel ID="plShowData" CssClass="ShowPanel" runat="server">
        <asp:GridView ID="gridAnecdoteCaseHistory" runat="server" AutoGenerateColumns="False"
            BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataKeyNames="HistoryNo" DataSourceID="sdsAnecdoteCaseHistoryList" GridLines="None" AllowPaging="True" PageSize="5" Width="100%">
            <Columns>
                <asp:CommandField ShowSelectButton="True" ButtonType="Button" />
                <asp:BoundField DataField="HistoryNo" HeaderText="HistoryNo" InsertVisible="False" ReadOnly="True" SortExpression="HistoryNo" Visible="False" />
                <asp:BoundField DataField="CaseNo" HeaderText="肇事單號" SortExpression="CaseNo" />
                <asp:CheckBoxField DataField="HasInsurance" HeaderText="出險" SortExpression="HasInsurance" Text="已出險" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="站別" SortExpression="DepName" />
                <asp:BoundField DataField="BuildDate" DataFormatString="{0:d}" HeaderText="發生日期" SortExpression="BuildDate" />
                <asp:BoundField DataField="BuildMan" HeaderText="BuildMan" SortExpression="BuildMan" Visible="False" />
                <asp:BoundField DataField="BuildMan_C" HeaderText="建檔人" SortExpression="BuildMan_C" />
                <asp:BoundField DataField="Car_ID" HeaderText="牌照號碼" SortExpression="Car_ID" />
                <asp:BoundField DataField="Driver" HeaderText="Driver" SortExpression="Driver" Visible="False" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員" SortExpression="DriverName" />
                <asp:BoundField DataField="InsuMan" HeaderText="保險經辦人" SortExpression="InsuMan" />
                <asp:TemplateField HeaderText="肇責比率" SortExpression="AnecdotalResRatio">
                    <ItemTemplate>
                        <asp:Label ID="eAnecdotalResRatio_Grid" runat="server" Text='<%# Eval("AnecdotalResRatio") %>' />
                        <asp:Label ID="lbPercent_Grid" runat="server" Text=" %" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:CheckBoxField DataField="IsNoDeduction" HeaderText="精勤扣款" SortExpression="IsNoDeduction" Text="免扣" />
                <asp:BoundField DataField="DeductionDate" HeaderText="扣發日期" SortExpression="DeductionDate" DataFormatString="{0:d}" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" Visible="false" />
                <asp:BoundField DataField="ModifyType" HeaderText="ModifyType" SortExpression="ModifyType" Visible="false" />
                <asp:BoundField DataField="ModifyType_C" HeaderText="異動類別" SortExpression="ModifyType_C" />
                <asp:BoundField DataField="ModifyDate" HeaderText="異動日期" SortExpression="ModifyDate" DataFormatString="{0:d}" />
                <asp:BoundField DataField="ModifyMan" HeaderText="ModifyMan" SortExpression="ModifyMan" Visible="False" />
                <asp:BoundField DataField="ModifyMan_C" HeaderText="異動人員" SortExpression="ModifyMan_C" />
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
        <asp:FormView ID="fvAnecdoteCaseHistory" runat="server" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataKeyNames="HistoryNo" DataSourceID="sdsAnecdoteCaseHistoryDetail" Width="100%">
            <ItemTemplate>
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseNo_List" runat="server" CssClass="text-Right-Blue" Text="肇事單號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModify_List" runat="server" CssClass="text-Right-Blue" Text="異動說明：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="5">
                            <asp:Label ID="eModifyType_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyType_C") %>' Width="15%" />
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="20%" />
                            <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="15%" />
                            <asp:Label ID="eModifyMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="35%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildDate_List" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuildDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuildMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="30%" />
                            <asp:Label ID="eBuildMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan_C") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbHasInsurance_List" runat="server" CssClass="text-Right-Blue" Text="出險狀況：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:CheckBox ID="HasInsuranceCheckBox" runat="server" CssClass="text-Left-Black" Checked='<%# Bind("HasInsurance") %>' Text="已出險" Enabled="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="30%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCar_ID_List" runat="server" CssClass="text-Right-Blue" Text="牌照號碼：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCar_ID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriver_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDriver_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver") %>' Width="30%" />
                            <asp:Label ID="eDriverName_LIst" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbInsuMan_List" runat="server" CssClass="text-Right-Blue" Text="保險經辦人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eInsuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InsuMan") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAnecdotalResRatio_List" runat="server" CssClass="text-Right-Blue" Text="肇責比率：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAnecdotalResRatio_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AnecdotalResRatio") %>' />
                            <asp:Label ID="lbPercent_List" runat="server" CssClass="text-Left-Black" Text=" %" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbIsNoDeduction_List" runat="server" CssClass="text-Right-Blue" Text="精勤扣款：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eIsNoDeduction_List" runat="server" CssClass="text-Left-Black" Checked='<%# Eval("IsNoDeduction") %>' Text="免扣" Enabled="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDeductionDate_List" runat="server" CssClass="text-Right-Blue" Text="扣發日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDeductionDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DeductionDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseOccurrence_List" runat="server" CssClass="text-Right-Blue" Text="肇事經過：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eCaseOccurrence_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("CaseOccurrence") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                            <asp:Label ID="eHistoryNo_List" runat="server" Text='<%# Eval("HistoryNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Height="95%" Width="97%" />
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
            <PagerStyle BackColor="#C6C3C6" ForeColor="Black" HorizontalAlign="Right" />
            <RowStyle BackColor="#DEDFDE" ForeColor="Black" />
        </asp:FormView>
        <asp:GridView ID="gridAnecdoteCaseBHistory" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BackColor="White" BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4"
            DataKeyNames="HistoryNoItems" DataSourceID="sdsAnecdoteCaseBHistoryList" PageSize="5" Width="100%">
            <Columns>
                <asp:CommandField ShowSelectButton="True" ButtonType="Button" />
                <asp:BoundField DataField="HistoryNo" HeaderText="HistoryNo" InsertVisible="False" ReadOnly="True" SortExpression="HistoryNo" Visible="False" />
                <asp:BoundField DataField="CaseNo" HeaderText="CaseNo" SortExpression="CaseNo" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="肇事單項次" SortExpression="Items" />
                <asp:BoundField DataField="CaseNoItems" HeaderText="CaseNoItems" SortExpression="CaseNoItems" Visible="False" />
                <asp:BoundField DataField="Relationship" HeaderText="對方姓名" SortExpression="Relationship" />
                <asp:BoundField DataField="RelCar_ID" HeaderText="對方車號" SortExpression="RelCar_ID" />
                <asp:BoundField DataField="EstimatedAmount" HeaderText="預估金額" SortExpression="EstimatedAmount" />
                <asp:BoundField DataField="ThirdInsurance" HeaderText="第三責任險" SortExpression="ThirdInsurance" />
                <asp:BoundField DataField="CompInsurance" HeaderText="強制險" SortExpression="CompInsurance" />
                <asp:BoundField DataField="PassengerInsu" HeaderText="乘客險" SortExpression="PassengerInsu" />
                <asp:BoundField DataField="DriverSharing" HeaderText="駕駛員負擔" SortExpression="DriverSharing" />
                <asp:BoundField DataField="CompanySharing" HeaderText="公司負擔" SortExpression="CompanySharing" />
                <asp:BoundField DataField="CarDamageAMT" HeaderText="車損金額" SortExpression="CarDamageAMT" />
                <asp:BoundField DataField="PersonDamageAMT" HeaderText="體傷金額" SortExpression="PersonDamageAMT" />
                <asp:BoundField DataField="RelationComp" HeaderText="對方賠付" SortExpression="RelationComp" />
                <asp:BoundField DataField="ReconciliationDate" HeaderText="和解日期" SortExpression="ReconciliationDate" DataFormatString="{0:d}" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" Visible="false" />
                <asp:BoundField DataField="ModifyType" HeaderText="ModifyType" SortExpression="ModifyType" Visible="false" />
                <asp:BoundField DataField="ModifyType_C" HeaderText="異動類別" SortExpression="ModifyType_C" />
                <asp:BoundField DataField="ModifyDate" HeaderText="異動日期" SortExpression="ModifyDate" DataFormatString="{0:d}" />
                <asp:BoundField DataField="ModifyMan" HeaderText="ModifyMan" SortExpression="ModifyMan" Visible="False" />
                <asp:BoundField DataField="ModifyMan_C" HeaderText="異動人員" SortExpression="ModifyMan_C" ReadOnly="True" />
            </Columns>
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="#FFFFCC" />
            <PagerStyle BackColor="#FFFFCC" ForeColor="#330099" HorizontalAlign="Center" />
            <RowStyle BackColor="White" ForeColor="#330099" />
            <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="#663399" />
            <SortedAscendingCellStyle BackColor="#FEFCEB" />
            <SortedAscendingHeaderStyle BackColor="#AF0101" />
            <SortedDescendingCellStyle BackColor="#F6F0C0" />
            <SortedDescendingHeaderStyle BackColor="#7E0000" />
        </asp:GridView>
        <asp:FormView ID="fvAnecdoteCaseBHistory" runat="server" DataKeyNames="HistoryNo" DataSourceID="sdsAnecdoteCaseBHistoryDetail" Width="100%">
            <ItemTemplate>
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseNo_ListB" runat="server" CssClass="text-Right-Blue" Text="肇事單號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseNo_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItems_ListB" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItems_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                            <asp:Label ID="eCaseNoItems_ListB" runat="server" Text='<%# Eval("CaseNoItems") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRelationShip_ListB" runat="server" CssClass="text-Right-Blue" Text="對方姓名：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eRelationship_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Relationship") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRelCar_ID_ListB" runat="server" CssClass="text-Right-Blue" Text="對方車號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eRelCar_ID_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelCar_ID") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbEstimatedAmount_ListB" runat="server" CssClass="text-Right-Blue" Text="預估金額" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eEstimatedAmount_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EstimatedAmount") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbThirdInsurance_ListB" runat="server" CssClass="text-Right-Blue" Text="第三責任險：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eThirdInsurance_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ThirdInsurance") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCompInsurance_ListB" runat="server" CssClass="text-Right-Blue" Text="強制險：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCompInsurance_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompInsurance") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPassengerInsu_ListB" runat="server" CssClass="text-Right-Blue" Text="乘客險：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="ePassengerInsu_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PassengerInsu") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCompanySharing_ListB" runat="server" CssClass="text-Right-Blue" Text="公司負擔：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCompanySharing_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompanySharing") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriverSharing_ListB" runat="server" CssClass="text-Right-Blue" Text="駕駛員負擔：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDriverSharing_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverSharing") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRelationComp_ListB" runat="server" CssClass="text-Right-Blue" Text="對方賠付：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eRelationComp_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelationComp") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCarDamageAMT_ListB" runat="server" CssClass="text-Right-Blue" Text="車損金額：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCarDamageAMT_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarDamageAMT") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPersonDamageAMT_ListB" runat="server" CssClass="text-Right-Blue" Text="體傷金額：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="ePersonDamageAMT_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PersonDamageAMT") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbReconciliationDate_ListB" runat="server" CssClass="text-Right-Blue" Text="和解日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eReconciliationDate_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ReconciliationDate","{0:yyyy/MM/dd}") %>' />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_ListB" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="5">
                            <asp:TextBox ID="eRemark_ListB" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModify_ListB" runat="server" CssClass="text-Right-Blue" Text="異動類別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyType_C_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyType_C") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_C_ListB" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyMan_C_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_ListB" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-8Col">
                            <asp:Label ID="eHistoryNo_ListB" runat="server" Text='<%# Eval("HistoryNo") %>' Visible="false" />
                        </td>
                        <td class="ColWidth-8Col" />
                        <td class="ColWidth-8Col" />
                        <td class="ColWidth-8Col" />
                        <td class="ColWidth-8Col" />
                        <td class="ColWidth-8Col" />
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsAnecdoteCaseHistoryList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT HistoryNo, CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuildMan)) AS BuildMan_C, Car_ID, Driver, DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, ModifyType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.ModifyType) AND (FKEY = CAST('客訴單歷史' AS char(16)) + CAST('fmServiceHistory' AS char(16)) + 'ModifyType')) AS ModifyType_C, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyMan_C, CaseOccurrence FROM AnecdoteCaseHistory AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsAnecdoteCaseHistoryDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT HistoryNo, CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuildMan)) AS BuildMan_C, Car_ID, Driver, DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, ModifyType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.ModifyType) AND (FKEY = CAST('客訴單歷史' AS char(16)) + CAST('fmServiceHistory' AS char(16)) + 'ModifyType')) AS ModifyType_C, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyMan_C, CaseOccurrence FROM AnecdoteCaseHistory AS a WHERE (HistoryNo = @HistoryNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridAnecdoteCaseHistory" Name="HistoryNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsAnecdoteCaseBHistoryList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT HistoryNo, CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, Remark, ModifyType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = b.ModifyType) AND (FKEY = CAST('客訴單歷史' AS char(16)) + CAST('fmServiceHistory' AS char(16)) + 'ModifyType')) AS ModifyType_C, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = b.ModifyMan)) AS ModifyMan_C, PassengerInsu, ItemsH, HistoryNoItems FROM AnecdoteCaseBHistory AS b WHERE (HistoryNo = @HistoryNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridAnecdoteCaseHistory" Name="HistoryNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsAnecdoteCaseBHistoryDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT HistoryNo, CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, Remark, ModifyType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = b.ModifyType) AND (FKEY = CAST('客訴單歷史' AS char(16)) + CAST('fmServiceHistory' AS char(16)) + 'ModifyType')) AS ModifyType_C, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = b.ModifyMan)) AS ModifyMan_C, PassengerInsu, ItemsH, HistoryNoItems FROM AnecdoteCaseBHistory AS b WHERE (HistoryNoItems = @HistoryNoItems)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridAnecdoteCaseBHistory" Name="HistoryNoItems" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
</asp:Content>
