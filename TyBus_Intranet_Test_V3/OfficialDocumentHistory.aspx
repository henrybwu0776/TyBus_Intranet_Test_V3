<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="OfficialDocumentHistory.aspx.cs" Inherits="TyBus_Intranet_Test_V3.OfficialDocumentHistory" %>

<asp:Content ID="OfficialDocumentHistoryForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">公文異動歷史記錄</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDocDep_Search" runat="server" CssClass="text-Right-Blue" Text="收發單位：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDocDep_Start_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDocDep_Start_Search_TextChanged" Width="40%" />
                    <asp:Label ID="lbSplit1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDocDep_End_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDocDep_End_Search_TextChanged" Width="40%" />
                    <br />
                    <asp:Label ID="eDocDepName_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit2" runat="server" CssClass="text-Left-Black" Width="5%" />
                    <asp:Label ID="eDocDepName_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepDate_Search" runat="server" CssClass="text-Right-Blue" Text="收發日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDocDate_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit3" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDocDate_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDocType_Search" runat="server" CssClass="text-Right-Blue" Text="文別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlDocType_Search" runat="server" CssClass="text-Left-Black" Width="90%"
                        DataSourceID="sdsDocType_Search" DataTextField="ClassTxt" DataValueField="ClassNo"
                        OnSelectedIndexChanged="ddlDocType_Search_SelectedIndexChanged" AutoPostBack="True">
                    </asp:DropDownList>
                    <br />
                    <asp:Label ID="eDocType_Search" runat="server" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbUndertaker_Search" runat="server" CssClass="text-Right-Blue" Text="承辦人：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eUndertaker_Start_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eUndertaker_Start_Search_TextChanged" Width="40%" />
                    <asp:Label ID="lbSplit4" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eUndertaker_End_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eUndertaker_End_Search_TextChanged" Width="40%" />
                    <br />
                    <asp:Label ID="eUndertakerName_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit5" runat="server" Width="5%" />
                    <asp:Label ID="eUndertakerName_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbModifyDate_Search" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eModifyDate_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit6" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eModifyDate_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbModifyMode_Search" runat="server" CssClass="text-Right-Blue" Text="異動類別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlModifyMode_Search" runat="server" CssClass="text-Left-Black" OnSelectedIndexChanged="ddlModifyMode_Search_SelectedIndexChanged" Width="90%">
                        <asp:ListItem Value="ALL" Text="全部" Selected="True" />
                        <asp:ListItem Value="DEL" Text="刪除" />
                        <asp:ListItem Value="EDIT" Text="修改" />
                    </asp:DropDownList>
                    <br />
                    <asp:Label ID="eModifyMode_Search" runat="server" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColWidth-8Col">
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
        <asp:GridView ID="gridOfficialDocumentHistoryList" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White"
            BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="DocIndex,Items,ModifyMode"
            DataSourceID="sdsOfficialDocumentHistoryList" GridLines="None" PageSize="5" Width="100%">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="DocIndex" HeaderText="序號" ReadOnly="True" SortExpression="DocIndex" />
                <asp:BoundField DataField="Items" HeaderText="項次" ReadOnly="True" SortExpression="Items" />
                <asp:BoundField DataField="DocDate" DataFormatString="{0:d}" HeaderText="收發日期" SortExpression="DocDate" />
                <asp:BoundField DataField="DocDep" HeaderText="DocDep" SortExpression="DocDep" Visible="False" />
                <asp:BoundField DataField="DocDepName" HeaderText="收發單位" ReadOnly="True" SortExpression="DocDepName" />
                <asp:BoundField DataField="DocFirstWord" HeaderText="DocFirstWord" SortExpression="DocFirstWord" Visible="False" />
                <asp:BoundField DataField="DocFirstWord_C" HeaderText="DocFirstWord_C" ReadOnly="True" SortExpression="DocFirstWord_C" Visible="False" />
                <asp:BoundField DataField="DocNo" HeaderText="DocNo" SortExpression="DocNo" Visible="False" />
                <asp:BoundField DataField="DocSourceUnit" HeaderText="來文機關" SortExpression="DocSourceUnit" />
                <asp:BoundField DataField="DocType" HeaderText="DocType" SortExpression="DocType" Visible="False" />
                <asp:BoundField DataField="DocType_C" HeaderText="文別" ReadOnly="True" SortExpression="DocType_C" />
                <asp:BoundField DataField="DocTitle" HeaderText="事由" SortExpression="DocTitle" />
                <asp:BoundField DataField="Undertaker" HeaderText="Undertaker" SortExpression="Undertaker" Visible="False" />
                <asp:BoundField DataField="UndertakerName" HeaderText="承辦人" ReadOnly="True" SortExpression="UndertakerName" />
                <asp:BoundField DataField="OutsideDocFirstWord" HeaderText="OutsideDocFirstWord" SortExpression="OutsideDocFirstWord" Visible="False" />
                <asp:BoundField DataField="OutsideDocNo" HeaderText="OutsideDocNo" SortExpression="OutsideDocNo" Visible="False" />
                <asp:BoundField DataField="Attachement" HeaderText="Attachement" SortExpression="Attachement" Visible="False" />
                <asp:BoundField DataField="Implementation" HeaderText="Implementation" SortExpression="Implementation" Visible="False" />
                <asp:BoundField DataField="BuildMan" HeaderText="BuildMan" SortExpression="BuildMan" Visible="False" />
                <asp:BoundField DataField="BuildManName" HeaderText="BuildManName" ReadOnly="True" SortExpression="BuildManName" Visible="False" />
                <asp:BoundField DataField="BuildDate" DataFormatString="{0:d}" HeaderText="BuildDate" SortExpression="BuildDate" Visible="False" />
                <asp:BoundField DataField="Remark" HeaderText="說明" SortExpression="Remark" />
                <asp:BoundField DataField="ModifyDate" DataFormatString="{0:d}" HeaderText="異動日期" SortExpression="ModifyDate" />
                <asp:BoundField DataField="ModifyMode" HeaderText="異動類別" ReadOnly="True" SortExpression="ModifyMode" />
                <asp:BoundField DataField="ModifyMan" HeaderText="ModifyMan" SortExpression="ModifyMan" Visible="False" />
                <asp:BoundField DataField="ModifyManName" HeaderText="異動人員" ReadOnly="True" SortExpression="ModifyManName" />
                <asp:BoundField DataField="DocYears" HeaderText="DocYears" SortExpression="DocYears" Visible="False" />
                <asp:CheckBoxField DataField="IsHide" HeaderText="IsHide" SortExpression="IsHide" Visible="False" />
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
        <asp:FormView ID="fvOfficialDocumentHistoryDetail" runat="server" DataKeyNames="DocIndex,Items,ModifyMode" DataSourceID="sdsOfficialDocumentHistoryDetail" Width="100%">
            <ItemTemplate>
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDocNo_List" runat="server" CssClass="text-Right-Blue" Text="收發文號：" Width="90%" />
                            <asp:Label ID="eDocINdex_List" runat="server" Text='<%# Eval("DocIndex") %>' Visible="false" />
                            <asp:Label ID="eDocFirstWord_List" runat="server" Text='<%# Eval("DocFirstWord") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDocYears_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocYears") %>' />
                            <asp:Label ID="lbSplit_List_1" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                            <asp:Label ID="eDocFirstWord_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocFirstWord_C") %>' />
                            <asp:Label ID="lbSplit_List_2" runat="server" CssClass="text-Left-Black" Text=" 字 第 " />
                            <asp:Label ID="eDocNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocNo") %>' />
                            <asp:Label ID="lbSplit_List_3" runat="server" CssClass="text-Left-Black" Text=" 號" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbItems_List" runat="server" CssClass="text-Right-Blue" Text="項次：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="30%" />
                            <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMode_List" runat="server" CssClass="text-Right-Blue" Text="異動類別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eModifyMode_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMode") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDocDep_List" runat="server" CssClass="text-Right-Blue" Text="收發文單位：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDocDep_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocDep") %>' Width="20%" />
                            <asp:Label ID="eDocDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocDepName") %>' Width="70%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDocDate_List" runat="server" CssClass="text-Right-Blue" Text="收發日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDocDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbUndertaker_List" runat="server" CssClass="text-Right-Blue" Text="承辦人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eUndertaker_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Undertaker") %>' Width="30%" />
                            <asp:Label ID="eUndertakerName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UndertakerName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDocType_List" runat="server" CssClass="text-Right-Blue" Text="文別：" Width="90%" />
                            <asp:Label ID="eDocType_List" runat="server" Text='<%# Eval("DocType") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDocType_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocType_C") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDocSourceUnit_List" runat="server" CssClass="text-Right-Blue" Text="來文機關：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDocSourceUnit_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocSourceUnit") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbOutsideDocNo_List" runat="server" CssClass="text-Right-Blue" Text="來文字號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eOutsideDocFirstWord_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OutsideDocFirstWord") %>' />
                            <asp:Label ID="lbSplit_List_4" runat="server" CssClass="text-Left-Black" Text=" 字 第 " />
                            <asp:Label ID="eOutsideDocNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OutsideDocNo") %>' />
                            <asp:Label ID="lbSplit_List_5" runat="server" CssClass="text-Left-Black" Text=" 號" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDocTitle_List" runat="server" CssClass="text-Right-Blue" Text="事由：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:Label ID="eDocTitle_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocTitle") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAttachment_List" runat="server" CssClass="text-Right-Blue" Text="收發文附件：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eAttachement_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Attachement") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbImplementation_List" runat="server" CssClass="text-Right-Blue" Text="辦理情形：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eImplementation_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Implementation") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuildMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="30%" />
                            <asp:Label ID="eBuildManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildManName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuildDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備考：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="RemarkLabel" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Remark") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStoreMan_List" runat="server" CssClass="text-Right-Blue" Text="歸檔人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStoreMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StoreMan") %>' Width="30%" />
                            <asp:Label ID="eStoreManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StoreManName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStoreDate_List" runat="server" CssClass="text-Right-Blue" Text="歸檔日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStoreDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StoreDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_Store_List" runat="server" CssClass="text-Right-Blue" Text="歸檔說明：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="Remark_StoreLabel" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Remark_Store") %>' Width="97%" />
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
    </asp:Panel>
    <asp:SqlDataSource ID="sdsOfficialDocumentHistoryList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT DocIndex, Items, DocDate, DocDep, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DocDep)) AS DocDepName, DocFirstWord, (SELECT DocFirstCWord FROM DOCFirstWord WHERE (FWNo = a.DocFirstWord)) AS DocFirstWord_C, DocNo, DocSourceUnit, DocType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '公文收發登記    OfficialDocumentDocType') AND (CLASSNO = a.DocType)) AS DocType_C, DocTitle, Undertaker, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.Undertaker)) AS UndertakerName, OutsideDocFirstWord, OutsideDocNo, Attachement, Implementation, BuildMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = a.BuildMan)) AS BuildManName, BuildDate, Remark, ModifyDate, ModifyMode, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName, DocYears, IsHide FROM OfficialDocumentHistory AS a WHERE (1 &lt;&gt; 1) ORDER BY DocDate DESC, DocIndex, Items"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsOfficialDocumentHistoryDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT DocIndex, Items, DocDate, DocDep, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DocDep)) AS DocDepName, DocFirstWord, (SELECT DocFirstCWord FROM DOCFirstWord WHERE (FWNo = a.DocFirstWord)) AS DocFirstWord_C, DocNo, DocSourceUnit, DocType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '公文收發登記    OfficialDocumentDocType') AND (CLASSNO = a.DocType)) AS DocType_C, DocTitle, Undertaker, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.Undertaker)) AS UndertakerName, OutsideDocFirstWord, OutsideDocNo, Attachement, Implementation, BuildMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (EMPNO = a.BuildMan)) AS BuildManName, BuildDate, Remark, StoreDate, StoreMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = a.StoreMan)) AS StoreManName, Remark_Store, ModifyDate, CASE WHEN isnull(ModifyMode , 'DEL') = 'DEL' THEN '刪除' ELSE '修改' END AS ModifyMode, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName, DocYears, IsHide FROM OfficialDocumentHistory AS a WHERE (DocIndex = @DocIndex) AND (Items = @Items) AND (ModifyMode = @ModifyMode)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridOfficialDocumentHistoryList" Name="DocIndex" PropertyName="SelectedDataKey[0]" />
            <asp:ControlParameter ControlID="gridOfficialDocumentHistoryList" Name="Items" PropertyName="SelectedDataKey[1]" />
            <asp:ControlParameter ControlID="gridOfficialDocumentHistoryList" Name="ModifyMode" PropertyName="SelectedDataKey[2]" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsDocType_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = CAST('公文收發登記' AS char(16)) + CAST('OfficialDocument' AS char(16)) + 'DocType') ORDER BY ClassNo"></asp:SqlDataSource>
</asp:Content>
