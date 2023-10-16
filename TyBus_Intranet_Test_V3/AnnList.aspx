<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="AnnList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.AnnList" %>

<asp:Content ID="AnnListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">公告管理</a>
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
    <asp:Panel ID="plShow" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridAnnList" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="AnnNo" DataSourceID="sdsAnnList_List"
            ForeColor="#333333" GridLines="None" PageSize="5" Width="100%" OnPageIndexChanging="gridAnnList_PageIndexChanging">
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
        <asp:SqlDataSource ID="sdsAnnList_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT AnnNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, EmpNo, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.EmpNo)) AS EmpName, PostDate, StartDate, EndDate, PostTitle FROM AnnList AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
        <asp:FormView ID="fvAnnList_Detail" runat="server" DataKeyNames="AnnNo" DataSourceID="sdsAnnList_Detail" OnDataBound="fvAnnList_Detail_DataBound" Width="100%">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CssClass="button-Blue" CausesValidation="True" CommandName="Update" Text="確定" Width="120px" UseSubmitBehavior="False" OnClientClick="this.disabled=true;" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAnnNo_Edit" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eAnnNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AnnNo") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="發文單位：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="30%" />
                                    <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbEmpNo_Edit" runat="server" CssClass="text-Right-Blue" Text="發文人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eEmpNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="30%" />
                                    <asp:Label ID="eEmpName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPostTitle_Edit" runat="server" CssClass="text-Right-Blue" Text="發文主旨：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="ePostTitle_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("PostTitle") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPostDate_Edit" runat="server" CssClass="text-Right-Blue" Text="發文日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="ePostDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PostDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbStartDate_Edit" runat="server" CssClass="text-Right-Blue" Text="有效日(起)：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eStartDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StartDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbEndDate_Edit" runat="server" CssClass="text-Right-Blue" Text="有效日(迄)：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eEndDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EndDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_Low ColBorder ColWidth-8Col" rowspan="2">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="公告本文：" Width="90%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-8Col" rowspan="2" colspan="5">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Height="98%" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="cbPostOpen_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbPostOpen_Edit_CheckedChanged" Text="發布在首頁" />
                                    <asp:TextBox ID="ePostOpen_Edit" runat="server" Text='<%#Bind("PostOpen") %>' Visible="false" />
                                    <asp:CheckBox ID="cbSendMail_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbSendMail_Edit_CheckedChanged" Text="發送 eMail (附檔上限 4MB)" />
                                    <asp:TextBox ID="eSendMail_Edit" runat="server" Text='<%# Bind("SendMail") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPostFiles_Edit" runat="server" CssClass="text-Right-Blue" Text="附檔檔案：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:FileUpload ID="fuPostFiles_Edit" runat="server" CssClass="text-Left-Black" Width="90%" />
                                    <asp:TextBox ID="ePostFiles_Edit" runat="server" Text='<%# Bind("PostFiles") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ModifyMan") %>' Width="30%" />
                                    <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CssClass="button-Blue" CausesValidation="True" OnClick="bbOK_INS_Click" Text="確定" Width="120px" UseSubmitBehavior="False" OnClientClick="this.disabled=true;" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAnnNo_INS" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eAnnNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AnnNo") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="發文單位：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Eval("DepNo") %>' Width="30%" />
                                    <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbEmpNo_INS" runat="server" CssClass="text-Right-Blue" Text="發文人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eEmpNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_INS_TextChanged" Text='<%# Eval("EmpNo") %>' Width="30%" />
                                    <asp:Label ID="eEmpName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPostTitle_INS" runat="server" CssClass="text-Right-Blue" Text="發文主旨：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="ePostTitle_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PostTitle") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPostDate_INS" runat="server" CssClass="text-Right-Blue" Text="發文日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="ePostDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PostDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbStartDate_INS" runat="server" CssClass="text-Right-Blue" Text="有效日(起)：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eStartDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StartDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbEndDate_INS" runat="server" CssClass="text-Right-Blue" Text="有效日(迄)：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eEndDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EndDate", "{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_Low ColBorder ColWidth-8Col" rowspan="2">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="公告本文：" Width="90%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-8Col" rowspan="2" colspan="5">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Height="98%" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:CheckBox ID="cbPostOpen_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbPostOpen_INS_CheckedChanged" Text="發布在首頁" />
                                    <asp:TextBox ID="ePostOpen_INS" runat="server" Text='<%#Eval("PostOpen") %>' Visible="false" />
                                    <asp:CheckBox ID="cbSendMail_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbSendMail_INS_CheckedChanged" Text="發送 eMail (附檔上限 4MB)" />
                                    <asp:TextBox ID="eSendMail_INS" runat="server" Text='<%# Eval("SendMail") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbPostFiles_INS" runat="server" CssClass="text-Right-Blue" Text="附檔檔名：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:FileUpload ID="fuPostFiles_INS" runat="server" CssClass="text-Left-Black" Width="90%" />
                                    <asp:TextBox ID="ePostFiles_INS" runat="server" Text='<%# Eval("PostFiles") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                                    <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CssClass="button-Black" CausesValidation="False" CommandName="New" Text="新增" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_LIst" runat="server" CssClass="button-Black" CausesValidation="False" CommandName="New" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CssClass="button-Blue" CausesValidation="False" CommandName="Edit" Text="編輯" Width="120px" />
                &nbsp;<asp:Button ID="bbDelete_List" runat="server" CssClass="button-Red" CausesValidation="False" OnClick="bbDelete_List_Click" Text="刪除" Width="120px" />
                &nbsp;<asp:Button ID="bbDownLoad_List" runat="server" CssClass="button-Blue" CausesValidation="false" OnClick="bbDownLoad_List_Click" Text="下載附檔" Width="120px" />
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
                            <asp:CheckBox ID="cbPostOpen_List" runat="server" CssClass="text-Left-Black" Text="發布在首頁" />
                            <asp:Label ID="ePostOpen_List" runat="server" Text='<%# Eval("PostOpen") %>' Visible="false" />
                            <asp:CheckBox ID="cbSendMail_List" runat="server" CssClass="text-Left-Black" Text="發送 eMail (附檔上限 4MB)" />
                            <asp:Label ID="eSendMail_List" runat="server" Text='<%# Eval("SendMail") %>' Visible="false" />
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

        <asp:SqlDataSource ID="sdsAnnList_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
            DeleteCommand="DELETE FROM AnnList WHERE (AnnNo = @AnnNo)"
            InsertCommand="INSERT INTO AnnList(AnnNo, DepNo, EmpNo, PostDate, StartDate, EndDate, PostTitle, Remark, PostFiles, BuMan, BuDate, PostOpen, SendMail) VALUES (@AnnNo, @DepNo, @EmpNo, @PostDate, @StartDate, @EndDate, @PostTitle, @Remark, @PostFiles, @BuMan, @BuDate, @PostOpen, @SendMail)"
            SelectCommand="SELECT AnnNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, EmpNo, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.EmpNo)) AS EmpName, PostDate, StartDate, EndDate, PostTitle, Remark, PostFiles, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = a.BuMan)) AS BuManName, BuDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName, ModifyDate, PostOpen, SendMail FROM AnnList AS a WHERE (AnnNo = @AnnNo)"
            UpdateCommand="UPDATE AnnList SET StartDate = @StartDate, EndDate = @EndDate, PostTitle = @PostTitle, Remark = @Remark, PostFiles = @PostFiles, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate, PostOpen = @PostOpen, PostDate = @PostDate, SendMail = @SendMail WHERE (AnnNo = @AnnNo)"
            OnDeleted="sdsAnnList_Detail_Deleted"
            OnInserted="sdsAnnList_Detail_Inserted"
            OnUpdated="sdsAnnList_Detail_Updated"
            OnUpdating="sdsAnnList_Detail_Updating">
            <DeleteParameters>
                <asp:Parameter Name="AnnNo" />
            </DeleteParameters>
            <InsertParameters>
                <asp:Parameter Name="AnnNo" />
                <asp:Parameter Name="DepNo" />
                <asp:Parameter Name="EmpNo" />
                <asp:Parameter Name="PostDate" />
                <asp:Parameter Name="StartDate" />
                <asp:Parameter Name="EndDate" />
                <asp:Parameter Name="PostTitle" />
                <asp:Parameter Name="Remark" />
                <asp:Parameter Name="PostFiles" />
                <asp:Parameter Name="BuMan" />
                <asp:Parameter Name="BuDate" />
                <asp:Parameter Name="PostOpen" />
                <asp:Parameter Name="SendMail" />
            </InsertParameters>
            <SelectParameters>
                <asp:ControlParameter ControlID="gridAnnList" Name="AnnNo" PropertyName="SelectedValue" />
            </SelectParameters>
            <UpdateParameters>
                <asp:Parameter Name="StartDate" />
                <asp:Parameter Name="EndDate" />
                <asp:Parameter Name="PostTitle" />
                <asp:Parameter Name="Remark" />
                <asp:Parameter Name="PostFiles" />
                <asp:Parameter Name="ModifyMan" />
                <asp:Parameter Name="ModifyDate" />
                <asp:Parameter Name="AnnNo" />
                <asp:Parameter Name="PostOpen" />
                <asp:Parameter Name="PostDate" />
                <asp:Parameter Name="SendMail" />
            </UpdateParameters>
        </asp:SqlDataSource>
    </asp:Panel>
</asp:Content>
