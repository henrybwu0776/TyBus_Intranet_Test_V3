<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmpMultiDep.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmpMultiDep" %>

<asp:Content ID="EmpMultiDepForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">員工多單位設定</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbEmpNo_Search" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eEmpNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_Search_TextChanged" Width="30%" />
                    <asp:Label ID="lbEmpName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Search_TextChanged" Width="30%" />
                    <asp:Label ID="lbDepName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbWorkType_Search" runat="server" CssClass="text-Right-Blue" Text="在職狀況：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlWorkType_Search" runat="server" CssClass="text-Left-Black" Width="90%" DataSourceID="sdsWorkTypeSearch" DataTextField="WorkType" DataValueField="WorkType">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
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
        <asp:GridView ID="gridEmpList" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px"
            CellPadding="4" DataKeyNames="EMPNO" DataSourceID="sdsEmpList" Width="100%" OnDataBound="gridEmpList_DataBound" OnSelectedIndexChanged="gridEmpList_SelectedIndexChanged">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="EMPNO" HeaderText="員工編號" ReadOnly="True" SortExpression="EMPNO" />
                <asp:BoundField DataField="NAME" HeaderText="姓名" SortExpression="NAME" />
                <asp:BoundField DataField="DEPNO" HeaderText="單位編號" SortExpression="DEPNO" />
                <asp:BoundField DataField="DepName" HeaderText="所屬單位" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="TYPE" HeaderText="TYPE" SortExpression="TYPE" Visible="False" />
                <asp:BoundField DataField="Type_C" HeaderText="員工類別" ReadOnly="True" SortExpression="Type_C" />
                <asp:BoundField DataField="worktype" HeaderText="在職狀況" SortExpression="worktype" />
                <asp:BoundField DataField="TITLE" HeaderText="TITLE" SortExpression="TITLE" Visible="False" />
                <asp:BoundField DataField="Title_C" HeaderText="職稱" ReadOnly="True" SortExpression="Title_C" />
            </Columns>
            <EmptyDataTemplate>
                <asp:Label ID="lbEmpDataList" runat="server" CssClass="errorMessageText" Text="請先查詢資料"></asp:Label>
            </EmptyDataTemplate>
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
        <asp:GridView ID="gridDepDetails" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#DEDFDE" BorderStyle="None" BorderWidth="1px" CellPadding="4"
            DataKeyNames="EmpNoItems" DataSourceID="sdsEmpDepDetailsList" ForeColor="Black" GridLines="Vertical" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="EmpNoItems" HeaderText="EmpNoItems" ReadOnly="True" SortExpression="EmpNoItems" Visible="False" />
                <asp:BoundField DataField="EmpNo" HeaderText="員工編號" SortExpression="EmpNo" />
                <asp:BoundField DataField="EmpNo_C" HeaderText="姓名" ReadOnly="True" SortExpression="EmpNo_C" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                <asp:BoundField DataField="DepNo" HeaderText="單位編號" SortExpression="DepNo" />
                <asp:BoundField DataField="DepNo_C" HeaderText="單位" ReadOnly="True" SortExpression="DepNo_C" />
                <asp:CheckBoxField DataField="IsDefault" HeaderText="是預設單位" SortExpression="IsDefault" />
                <asp:CheckBoxField DataField="IsUsed" HeaderText="使用中" SortExpression="IsUsed" />
                <asp:BoundField DataField="Remark" HeaderText="Remark" SortExpression="Remark" Visible="False" />
                <asp:BoundField DataField="BuMan" HeaderText="BuMan" SortExpression="BuMan" Visible="False" />
                <asp:BoundField DataField="BuMan_C" HeaderText="BuMan_C" ReadOnly="True" SortExpression="BuMan_C" Visible="False" />
                <asp:BoundField DataField="BuDate" HeaderText="BuDate" SortExpression="BuDate" Visible="False" />
            </Columns>
            <EmptyDataTemplate>
                <asp:Label ID="lbEmptyDetail" runat="server" CssClass="errorMessageText" Text="查無所屬站別資料"></asp:Label>
            </EmptyDataTemplate>
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
        <asp:FormView ID="fvEmpDepDetails" runat="server" DataKeyNames="EmpNoItems" DataSourceID="sdsEmpDepDetails" Width="100%" OnDataBound="fvEmpDepDetails_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Update" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Update" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Update" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbEmpNo_Edit" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="90%" />
                                <asp:Label ID="eEmpNoItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNoItems") %>' Visible="false" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                <asp:TextBox ID="eEmpNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Enabled="false" Width="30%" />
                                <asp:Label ID="eEmpNo_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo_C") %>' Width="55%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbItems_Edit" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="eItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                            </td>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Text='<%# Bind("DepNo") %>' Width="30%" />
                                    <asp:Label ID="eDepNo_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo_C") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eIsDefault_Edit" runat="server" CssClass="text-Left-Black" Checked='<%# Bind("IsDefault") %>' Text="是主要所屬單位" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eIsUsed_Edit" runat="server" CssClass="text-Left-Black" Checked='<%# Bind("IsUsed") %>' Text="使用中" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="4">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Height="95%" Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                                    <asp:Label ID="eBuMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_Insert" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Insert" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Insert" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbEmpNo_INS" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="90%" />
                                <asp:Label ID="eEmpNoItems_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNoItems") %>' Visible="false" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                <asp:TextBox ID="eEmpNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_INS_TextChanged" Text='<%# Bind("EmpNo") %>' Width="30%" />
                                <asp:Label ID="eEmpNo_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo_C") %>' Width="55%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbItems_INS" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="eItems_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                            </td>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Bind("DepNo") %>' Width="30%" />
                                    <asp:Label ID="eDepNo_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo_C") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eIsDefault_INS" runat="server" CssClass="text-Left-Black" Checked='<%# Bind("IsDefault") %>' Text="是主要所屬單位" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eIsUsed_INS" runat="server" CssClass="text-Left-Black" Checked='<%# Bind("IsUsed") %>' Text="使用中" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="4">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Height="95%" Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                                    <asp:Label ID="eBuMan_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="Edit" Text="編輯" Width="120px" />
                &nbsp;<asp:Button ID="bbDelete" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Delete" Text="刪除" Width="120px" />
                <table class="TableSetting">
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbEmpNo_List" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="90%" />
                        <asp:Label ID="eEmpNoItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNoItems") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                        <asp:Label ID="eEmpNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="30%" />
                        <asp:Label ID="eEmpNo_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo_C") %>' Width="55%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbItems_List" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                    </td>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="30%" />
                            <asp:Label ID="eDepNo_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo_C") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eIsDefault_List" runat="server" CssClass="text-Left-Black" Checked='<%# Eval("IsDefault") %>' Enabled="false" Text="是主要所屬單位" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eIsUsed_List" runat="server" CssClass="text-Left-Black" Checked='<%# Eval("IsUsed") %>' Enabled="false" Text="使用中" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="4">
                            <asp:Label ID="eRemark_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Remark") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                            <asp:Label ID="eBuMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
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
    <asp:SqlDataSource ID="sdsEmpList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT EMPNO, NAME, DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DEPNO)) AS DepName, TYPE, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '人事資料檔      EMPLOYEE        type') AND (CLASSNO = e.TYPE)) AS Type_C, worktype, TITLE, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (FKEY = '員工資料        EMPLOYEE        TITLE') AND (CLASSNO = e.TITLE)) AS Title_C FROM EMPLOYEE AS e WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsEmpDepDetailsList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT EmpNoItems, EmpNo, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.EmpNo)) AS EmpNo_C, Items, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepNo_C, IsDefault, IsUsed, Remark, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.BuMan)) AS BuMan_C, BuDate FROM EmployeeDepNo AS a WHERE (EmpNo = @EmpNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridEmpList" Name="EmpNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsEmpDepDetails" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        DeleteCommand="DELETE FROM EmployeeDepNo WHERE (EmpNoItems = @EmpNoItems)"
        InsertCommand="INSERT INTO EmployeeDepNo(EmpNoItems, EmpNo, Items, DepNo, IsDefault, IsUsed, Remark, BuMan, BuDate) VALUES (@EmpNoItems, @EmpNo, @Items, @DepNo, @IsDefault, @IsUsed, @Remark, @BuMan, @BuDate)"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT EmpNoItems, EmpNo, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.EmpNo)) AS EmpNo_C, Items, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepNo_C, IsDefault, IsUsed, Remark, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.BuMan)) AS BuMan_C, BuDate FROM EmployeeDepNo AS a WHERE (EmpNoItems = @EmpNoItems)"
        UpdateCommand="UPDATE EmployeeDepNo SET DepNo = @DepNo, IsDefault = @IsDefault, IsUsed = @IsUsed, Remark = @Remark WHERE (EmpNoItems = @EmpNoItems)"
        OnDeleted="sdsEmpDepDetails_Deleted"
        OnInserted="sdsEmpDepDetails_Inserted"
        OnInserting="sdsEmpDepDetails_Inserting"
        OnUpdated="sdsEmpDepDetails_Updated">
        <DeleteParameters>
            <asp:Parameter Name="EmpNoItems" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="EmpNoItems" />
            <asp:Parameter Name="EmpNo" />
            <asp:Parameter Name="Items" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="IsDefault" />
            <asp:Parameter Name="IsUsed" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="BuDate" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridDepDetails" Name="EmpNoItems" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="IsDefault" />
            <asp:Parameter Name="IsUsed" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="EmpNoItems" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsWorkTypeSearch" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST('' AS varchar) AS WorkType UNION ALL SELECT DISTINCT worktype FROM EMPLOYEE WHERE (ISNULL(worktype, '') &lt;&gt; '') ORDER BY worktype"></asp:SqlDataSource>
</asp:Content>
