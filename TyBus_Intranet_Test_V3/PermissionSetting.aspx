<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="PermissionSetting.aspx.cs" Inherits="TyBus_Intranet_Test_V3.PermissionSetting" %>

<asp:Content ID="PermissionSettingForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">網頁權限設定</a>
    </div>
    <br />
    <asp:Panel ID="plSearech" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColWidth-6Col ColHeight ColBorder">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="單位別：" Width="90%" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder" colspan="2">
                    <asp:DropDownList ID="ddlDepNo_Start" runat="server" CssClass="text-Left-Black" Width="45%"
                        DataSourceID="sdsDepList_Start" DataTextField="NAME" DataValueField="DEPNO"
                        OnSelectedIndexChanged="ddlDepNo_Start_SelectedIndexChanged" />
                    <asp:Label ID="lbSplit_DepNo" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:DropDownList ID="ddlDepNo_End" runat="server" CssClass="text-Left-Black" Width="45%"
                        DataSourceID="sdsDepList_End" DataTextField="NAME" DataValueField="DEPNO"
                        OnSelectedIndexChanged="ddlDepNo_End_SelectedIndexChanged" />
                    <br />
                    <asp:Label ID="eDepNo_Start_Search" runat="server" Visible="false" />
                    <asp:Label ID="eDepNo_End_Search" runat="server" Visible="false" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder">
                    <asp:Label ID="lbEmpNoStart_Search" runat="server" CssClass="text-Right-Blue" Text="員工別：" Width="90%" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder" colspan="2">
                    <asp:TextBox ID="eEmpNo_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_EmpNo" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eEmpNo_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col ColHeight ColBorder">
                    <asp:Label ID="lbGroupID_Search" runat="server" CssClass="text-Right-Blue" Text="所屬群組：" Width="90%" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder" colspan="2">
                    <asp:DropDownList ID="ddlGroupID_Search" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlGroupID_Search_SelectedIndexChanged" CssClass="text-Left-Black" Width="90%" DataSourceID="sdsGroupID_Search" DataTextField="CLASSTXT" DataValueField="CLASSNO"></asp:DropDownList>
                    <asp:TextBox ID="eGroupID_Search" runat="server" Visible="false" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder">
                    <asp:Label ID="lbControlCName_Search" runat="server" CssClass="text-Right-Blue" Text="功能名稱：" Width="90%" />
                </td>
                <td class="ColWidth-6Col ColHeight ColBorder" colspan="2">
                    <asp:TextBox ID="eControlCName_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" colspan="3">
                    <asp:FileUpload ID="fuExcel" runat="server" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbImportRight" runat="server" CssClass="button-Blue" Text="EXCEL 批次修改權限" OnClick="bbImportRight_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClearRight" runat="server" CssClass="button-Red" Text="解除個人權限" OnClick="bbClearRight_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" colspan="3" />
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col"></td>
                <td class="ColWidth-6Col"></td>
                <td class="ColWidth-6Col"></td>
                <td class="ColWidth-6Col"></td>
                <td class="ColWidth-6Col"></td>
                <td class="ColWidth-6Col"></td>
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plPermissionAData_Show" runat="server" CssClass="PanelMargin">
        <asp:GridView ID="gridWebPermissionA_List" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="ControlName"
            DataSourceID="sdsWebPermissionA_List" GridLines="None" PageSize="5" Width="100%">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="ControlName" HeaderText="功能名稱" SortExpression="ControlName" />
                <asp:BoundField DataField="GroupName" HeaderText="所屬群組" />
                <asp:BoundField DataField="ControlCName" HeaderText="功能中文名" SortExpression="ControlCName" />
                <asp:BoundField DataField="TargetPage" HeaderText="目標網頁" SortExpression="TargetPage" />
                <asp:BoundField DataField="Remark" HeaderText="備註說明" SortExpression="Remark" />
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
    <asp:Panel ID="plPermissionA_Detail" runat="server" CssClass="PanelMargin">
        <asp:FormView ID="fvPermissionA_Data" runat="server" DataKeyNames="ControlName" DataSourceID="sdsWebPermissionA_Data" Width="100%" OnDataBound="fvPermissionA_Data_DataBound">
            <EditItemTemplate>
                <asp:Button ID="UpdateButton_EditA" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Update" Text="確定" />
                &nbsp;
                <asp:Button ID="UpdateCancelButton_EditA" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" />
                <br />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbControlName_EditA" runat="server" CssClass="text-Right-Blue" Text="選單功能：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eControlName_EditA" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ControlName") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbGroupID_EditA" runat="server" CssClass="text-Right-Blue" Text="所屬群組：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:DropDownList ID="ddlGroupID_EditA" runat="server" CssClass="text-Left-Black" AutoPostBack="True" OnSelectedIndexChanged="ddlGroupID_EditA_SelectedIndexChanged" Width="90%" DataSourceID="sdsGroupID_Edit" DataTextField="CLASSTXT" DataValueField="CLASSNO" />
                            <asp:TextBox ID="eGroupID_EditA" runat="server" CssClass="text-Left-Black" Text='<%# Bind("GroupID") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbControlCName_EditA" runat="server" CssClass="text-Right-Blue" Text="選單中文名稱：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eControlCName_EditA" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ControlCName") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbTargetPage_EditA" runat="server" CssClass="text-Right-Blue" Text="繫結網頁：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eTargetPage_EditA" runat="server" CssClass="text-Left-Black" Text='<%# Bind("TargetPage") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_EditA" runat="server" CssClass="text-Right-Blue" Text="備註說明：" Width="90%" />
                            <asp:Label ID="eOrderIndex_EditA" runat="server" Text='<%# Bind("OrderIndex") %>' Visible="false" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="7">
                            <asp:TextBox ID="eRemark_EditA" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Width="90%" Height="95%" />
                        </td>
                    </tr>
                </table>
            </EditItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CssClass="button-Black" CommandName="New" Text="新增" />
            </EmptyDataTemplate>
            <InsertItemTemplate>
                <asp:Button ID="InsertButton_InsertA" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Insert" Text="確定" />
                &nbsp;
                <asp:Button ID="InsertCancelButton_InsertA" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" />
                <br />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbControlName_InsertA" runat="server" CssClass="text-Right-Blue" Text="選單功能：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eControlName_InsertA" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ControlName") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbGroupID_InsertA" runat="server" CssClass="text-Right-Blue" Text="所屬群組：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:DropDownList ID="ddlGroupID_InsertA" runat="server" CssClass="text-Left-Black" AutoPostBack="True" OnSelectedIndexChanged="ddlGroupID_InsertA_SelectedIndexChanged" Width="90%" DataSourceID="sdsGroupID_INS" DataTextField="CLASSTXT" DataValueField="CLASSNO"></asp:DropDownList>
                            <asp:TextBox ID="eGroupID_InsertA" runat="server" CssClass="text-Left-Black" Text='<%# Bind("GroupID") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbControlCName_InsertA" runat="server" CssClass="text-Right-Blue" Text="選單中文名稱：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eControlCName_InsertA" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ControlCName") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbTargetPage_InsertA" runat="server" CssClass="text-Right-Blue" Text="繫結網頁：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eTargetPage_InsertA" runat="server" CssClass="text-Left-Black" Text='<%# Bind("TargetPage") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_InsertA" runat="server" CssClass="text-Right-Blue" Text="備註說明：" Width="90%" />
                            <asp:Label ID="eOrderIndex_InsertA" runat="server" Text='<%# Eval("OrderIndex") %>' Visible="false" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="7">
                            <asp:TextBox ID="eRemark_InsertA" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Width="90%" Height="95%" />
                        </td>
                    </tr>
                </table>
            </InsertItemTemplate>
            <ItemTemplate>
                <asp:Button ID="NewButton_ListA" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" />
                &nbsp;
                <asp:Button ID="EditButton_ListA" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Edit" Text="修改" />
                &nbsp;
                <asp:Button ID="DeleteButton_ListA" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Delete" Text="刪除" />
                <br />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbControlName_ListA" runat="server" CssClass="text-Right-Blue" Text="選單功能：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eControlName_ListA" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ControlName") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbGroupID_ListA" runat="server" CssClass="text-Right-Blue" Text="所屬群組：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eGroupName_ListA" runat="server" CssClass="text-Left-Black" Text='<%# Eval("GroupName") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbControlCName_ListA" runat="server" CssClass="text-Right-Blue" Text="選單中文名稱：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eControlCName_ListA" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ControlCName") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbTargetPage_ListA" runat="server" CssClass="text-Right-Blue" Text="繫結網頁：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eTargetPage_ListA" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TargetPage") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_ListA" runat="server" CssClass="text-Right-Blue" Text="備註說明：" Width="90%" />
                            <asp:Label ID="eOrderIndex_ListA" runat="server" Text='<%# Eval("OrderIndex") %>' Visible="false" />
                        </td>
                        <td class="ColBorder ColWidth-8Col MultiLine-Low" colspan="7">
                            <asp:TextBox ID="eRemark_ListA" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" ReadOnly="true" Text='<%# Eval("Remark") %>' Width="90%" Height="95%" />
                        </td>
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:Panel ID="plPermissionBData_Show" runat="server" CssClass="PanelMargin">
        <asp:GridView ID="gridWebPermissionB_List" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BackColor="White" BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" DataKeyNames="ControlNameItems"
            DataSourceID="sdsWebPermissionB_List" PageSize="5" Width="100%">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="ControlName" HeaderText="ControlName" SortExpression="ControlName" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                <asp:BoundField DataField="ControlNameItems" HeaderText="ControlNameItems" ReadOnly="True" SortExpression="ControlNameItems" Visible="False" />
                <asp:BoundField DataField="DepNo" HeaderText="單位代號" SortExpression="DepNo" />
                <asp:BoundField DataField="DepName" HeaderText="單位" SortExpression="DepName" />
                <asp:BoundField DataField="EmpNo" HeaderText="員工工號" SortExpression="EmpNo" />
                <asp:BoundField DataField="EmpName" HeaderText="員工姓名" SortExpression="EmpName" />
                <asp:CheckBoxField DataField="AllowPermission" HeaderText="授與權限" SortExpression="AllowPermission" />
                <asp:BoundField DataField="RemarkB" HeaderText="附註" SortExpression="RemarkB" />
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
    </asp:Panel>
    <asp:Panel ID="plPermissionB_Detail" runat="server" CssClass="PanelMargin">
        <asp:FormView ID="fvPermissionB_Data" runat="server" DataKeyNames="ControlNameItems" DataSourceID="sdsWebPermissionB_Data" Width="100%">
            <EditItemTemplate>
                <asp:Button ID="UpdateButton_UpdateB" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Update" Text="更新" />
                &nbsp;
                <asp:Button ID="UpdateCancelButton_UpdateB" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" />
                <br />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbControlName_EditB" runat="server" CssClass="text-Right-Blue" Text="功能名稱：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eControlName_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ControlName") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItems_EditB" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItems_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                            <br />
                            <asp:Label ID="eControlNameItems_EditB" runat="server" Text='<%# Eval("ControlNameItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eAllowPermission_EditB" runat="server" CssClass="text-Left-Black" Text="授與權限" Checked='<%# Bind("AllowPermission") %>' Width="90%" />
                        </td>
                        <td class="ColWidth-8Col" />
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_EditB" runat="server" CssClass="text-Right-Blue" Text="單位別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="eDepNo_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DepNo") %>' AutoPostBack="true" OnTextChanged="eDepNo_EditB_TextChanged" Width="35%" />
                            <asp:Label ID="eDepName_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbEmpNo_EditB" runat="server" CssClass="text-Right-Blue" Text="員工別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="eEmpNo_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Bind("EmpNo") %>' AutoPostBack="true" OnTextChanged="eEmpNo_EditB_TextChanged" Width="35%" />
                            <asp:Label ID="eEmpName_EditB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemarkB_EditB" runat="server" CssClass="text-Right-Blue" Text="附註：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="5">
                            <asp:TextBox ID="eRemarkB_EditB" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("RemarkB") %>' Width="90%" Height="97%" />
                        </td>
                    </tr>
                </table>
            </EditItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="NewButton_EmptyB" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增明細" />
                <asp:Button ID="BatchNewButton_EmptyB" runat="server" CausesValidation="false" CssClass="button-Red" Text="批次新增明細" OnClick="BatchNewButton_EmptyB_Click" />
            </EmptyDataTemplate>
            <InsertItemTemplate>
                <asp:Button ID="InsertButton_InsertB" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Insert" Text="確定" />
                &nbsp;
                <asp:Button ID="InsertCancelButton_InsertB" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" />
                <br />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbControlName_InsertB" runat="server" CssClass="text-Right-Blue" Text="功能名稱：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eControlName_InsertB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ControlName") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItems_InsertB" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItems_InsertB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                            <br />
                            <asp:Label ID="eControlNameItems_InsertB" runat="server" Text='<%# Eval("ControlNameItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eAllowPermission_InsertB" runat="server" CssClass="text-Left-Black" Text="授與權限" Checked='<%# Bind("AllowPermission") %>' Width="90%" />
                        </td>
                        <td class="ColWidth-8Col" />
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_InsertB" runat="server" CssClass="text-Right-Blue" Text="單位別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="eDepNo_InsertB" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DepNo") %>' AutoPostBack="true" OnTextChanged="eDepNo_InsertB_TextChanged" Width="35%" />
                            <asp:Label ID="eDepName_InsertB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbEmpNo_InsertB" runat="server" CssClass="text-Right-Blue" Text="員工別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="eEmpNo_InsertB" runat="server" CssClass="text-Left-Black" Text='<%# Bind("EmpNo") %>' AutoPostBack="true" OnTextChanged="eEmpNo_InsertB_TextChanged" Width="35%" />
                            <asp:Label ID="eEmpName_InsertB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemarkB_InsertB" runat="server" CssClass="text-Right-Blue" Text="附註：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="5">
                            <asp:TextBox ID="eRemarkB_InsertB" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("RemarkB") %>' Width="90%" Height="97%" />
                        </td>
                    </tr>
                </table>
            </InsertItemTemplate>
            <ItemTemplate>
                <asp:Button ID="NewButton_ListB" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增明細" />
                &nbsp;
                <asp:Button ID="BatchNewButton_ListB" runat="server" CausesValidation="false" CssClass="button-Red" Text="批次新增明細" OnClick="BatchNewButton_EmptyB_Click" />
                &nbsp;
                <asp:Button ID="EditButton_ListB" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="Edit" Text="修改明細" />
                &nbsp;
                <asp:Button ID="DeleteButton_ListB" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Delete" Text="刪除明細" />
                <br />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbControlName_ListB" runat="server" CssClass="text-Right-Blue" Text="功能名稱：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eControlName_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ControlName") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItems_ListB" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eItems_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                            <br />
                            <asp:Label ID="eControlNameItems_ListB" runat="server" Text='<%# Eval("ControlNameItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eAllowPermission_ListB" runat="server" CssClass="text-Left-Black" Enabled="false" Text="授與權限" Checked='<%# Eval("AllowPermission") %>' Width="90%" />
                        </td>
                        <td class="ColWidth-8Col" />
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_ListB" runat="server" CssClass="text-Right-Blue" Text="單位別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbEmpNo_ListB" runat="server" CssClass="text-Right-Blue" Text="員工別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eEmpNo_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="35%" />
                            <asp:Label ID="eEmpName_ListB" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemarkB_ListB" runat="server" CssClass="text-Right-Blue" Text="附註：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="5">
                            <asp:TextBox ID="eRemarkB_ListB" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" ReadOnly="true" Text='<%# Eval("RemarkB") %>' Width="90%" Height="97%" />
                        </td>
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsWebPermissionA_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT ControlName, OrderIndex, ControlCName, TargetPage, Remark, GroupID, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.GroupID) AND (FKEY = '網頁功能權限群組WebPermission   GroupID')) AS GroupName FROM WebPermissionA AS a ORDER BY ControlName"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsWebPermissionB_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT ControlName, Items, ControlNameItems, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = WebPermissionB.DepNo)) AS DepName, EmpNo, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = WebPermissionB.EmpNo)) AS EmpName, AllowPermission, RemarkB FROM WebPermissionB WHERE (ControlName = @ControlName)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridWebPermissionA_List" Name="ControlName" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsDepList_Start" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CAST('' AS nvarchar) AS DEPNO, CAST('' AS nvarchar) AS NAME UNION ALL SELECT DEPNO, NAME FROM Department WHERE (DEPNO &gt;= '01') AND (DEPNO &lt;= '90')"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsDepList_End" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CAST('' AS nvarchar) AS DEPNO, CAST('' AS nvarchar) AS NAME UNION ALL SELECT DEPNO, NAME FROM Department WHERE (DEPNO &gt;= '01') AND (DEPNO &lt;= '90')"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsWebPermissionA_Data" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        DeleteCommand="DELETE FROM WebPermissionA WHERE (ControlName = @ControlName)"
        InsertCommand="INSERT INTO WebPermissionA(ControlName, OrderIndex, ControlCName, TargetPage, Remark, GroupID) VALUES (@ControlName, @OrderIndex, @ControlCName, @TargetPage, @Remark, @GroupID)"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT ControlName, OrderIndex, ControlCName, Remark, TargetPage, GroupID, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.GroupID) AND (FKEY = '網頁功能權限群組WebPermission   GroupID')) AS GroupName FROM WebPermissionA AS a WHERE (ControlName = @ControlName)"
        UpdateCommand="UPDATE WebPermissionA SET OrderIndex = @OrderIndex, ControlCName = @ControlCName, Remark = @Remark, TargetPage = @TargetPage, GroupID = @GroupID WHERE (ControlName = @ControlName)"
        OnDeleted="sdsWebPermissionA_Data_Deleted"
        OnInserted="sdsWebPermissionA_Data_Inserted"
        OnUpdated="sdsWebPermissionA_Data_Updated"
        OnInserting="sdsWebPermissionA_Data_Inserting">
        <DeleteParameters>
            <asp:Parameter Name="ControlName" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="ControlName" />
            <asp:Parameter Name="ControlCName" />
            <asp:Parameter Name="TargetPage" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="OrderIndex" />
            <asp:Parameter Name="GroupID" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridWebPermissionA_List" Name="ControlName" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="ControlName" />
            <asp:Parameter Name="ControlCName" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="TargetPage" />
            <asp:Parameter Name="OrderIndex" />
            <asp:Parameter Name="GroupID" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsWebPermissionB_Data" runat="server"
        DeleteCommand="DELETE FROM WebPermissionB WHERE (ControlNameItems = @ControlNameItems)"
        InsertCommand="INSERT INTO WebPermissionB(ControlName, Items, ControlNameItems, DepNo, EmpNo, AllowPermission, RemarkB) VALUES (@ControlName, @Items, @ControlNameItems, @DepNo, @EmpNo, @AllowPermission, @RemarkB)"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT ControlName, Items, ControlNameItems, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = WebPermissionB.DepNo)) AS DepName, EmpNo, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = WebPermissionB.EmpNo)) AS EmpName, AllowPermission, RemarkB FROM WebPermissionB WHERE (ControlNameItems = @ControlNameItems)" UpdateCommand="UPDATE WebPermissionB SET DepNo = @DepNo, EmpNo = @EmpNo, AllowPermission = @AllowPermission, RemarkB = @RemarkB WHERE (ControlNameItems = @ControlNameItems)"
        OnDeleted="sdsWebPermissionB_Data_Deleted"
        OnInserted="sdsWebPermissionB_Data_Inserted"
        OnInserting="sdsWebPermissionB_Data_Inserting"
        OnUpdated="sdsWebPermissionB_Data_Updated" ConnectionString="<%$ ConnectionStrings:connERPSQL %>">
        <DeleteParameters>
            <asp:Parameter Name="ControlNameItems" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="ControlName" />
            <asp:Parameter Name="Items" />
            <asp:Parameter Name="ControlNameItems" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="EmpNo" />
            <asp:Parameter Name="AllowPermission" />
            <asp:Parameter Name="RemarkB" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridWebPermissionB_List" Name="ControlNameItems" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="EmpNo" />
            <asp:Parameter Name="AllowPermission" />
            <asp:Parameter Name="RemarkB" />
            <asp:Parameter Name="ControlNameItems" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsGroupID_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CAST(NULL AS varchar) AS ClassNo, CAST(NULL AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '網頁功能權限群組WebPermission   GroupID') ORDER BY CLASSNO"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsGroupID_INS" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '網頁功能權限群組WebPermission   GroupID') ORDER BY CLASSNO"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsGroupID_Edit" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '網頁功能權限群組WebPermission   GroupID') ORDER BY CLASSNO"></asp:SqlDataSource>
</asp:Content>
