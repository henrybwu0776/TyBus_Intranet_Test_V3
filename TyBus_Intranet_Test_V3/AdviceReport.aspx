<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="AdviceReport.aspx.cs" Inherits="TyBus_Intranet_Test_V3.AdviceReport" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="AdviceReportForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">員工疏失勸導單</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCaseDate_Search" runat="server" CssClass="text-Right-Blue" Text="事件日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eCaseDate_S_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eCaseDate_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="所屬單位:" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_S_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_S_Search_TextChanged" Width="40%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDepNo_E_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_E_Search_TextChanged" Width="40%" />
                    <br />
                    <asp:Label ID="eDepName_S_Search" runat="server" CssClass="text-Left-Black" Width="45%" />
                    <asp:Label ID="eDepName_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbEmpNo_Search" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eEmpNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eEmpName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" OnClick="bbSearch_Click" Text="搜尋" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" OnClick="bbExcel_Click" Text="匯出EXCEL" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" OnClick="bbClose_Click" Text="結束" Width="90%" />
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
        <asp:GridView ID="gridAdviceReportList" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False"
            BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" DataKeyNames="CaseNo"
            DataSourceID="sdsAdviceReport_List" GridLines="None" OnPageIndexChanging="gridAdviceReportList_PageIndexChanging" PageSize="5">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="CaseNo" HeaderText="單號" ReadOnly="True" SortExpression="CaseNo" />
                <asp:BoundField DataField="DepName" HeaderText="單位" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="EmpName" HeaderText="員工姓名" SortExpression="EmpName" />
                <asp:BoundField DataField="Title_C" HeaderText="職稱" ReadOnly="True" SortExpression="Title_C" />
                <asp:BoundField DataField="CaseDate" DataFormatString="{0:D}" HeaderText="事件日期" SortExpression="CaseDate" />
                <asp:BoundField DataField="CaseTime" HeaderText="事件時間" SortExpression="CaseTime" />
                <asp:BoundField DataField="Position" HeaderText="地點" SortExpression="Position" />
                <asp:BoundField DataField="CaseNote" HeaderText="勸導事由" SortExpression="CaseNote" />
                <asp:BoundField DataField="Remark" HeaderText="補充說明" SortExpression="Remark" />
                <asp:BoundField DataField="DenialReason" HeaderText="拒簽原因" SortExpression="DenialReason" />
                <asp:BoundField DataField="BuDate" DataFormatString="{0:D}" HeaderText="開單日期" SortExpression="BuDate" />
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
        <asp:SqlDataSource ID="sdsAdviceReport_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
            SelectCommand="SELECT CaseNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, EmpNo, EmpName, Title, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '員工資料        EMPLOYEE        TITLE') AND (CLASSNO = a.Title)) AS Title_C, CaseDate, CaseTime, Position, CaseNote, Remark, DenialReason, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuManName, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName, SourceType FROM AdviceReport AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    </asp:Panel>
    <asp:FormView ID="fvAdviceReportDetail" runat="server" DataKeyNames="CaseNo" DataSourceID="sdsAdviceReport_Detail" Width="100%"
        OnDataBound="fvAdviceReportDetail_DataBound">
        <EditItemTemplate>
            <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" OnClick="bbOK_Edit_Click" CssClass="button-Black" Text="確定" Width="120px" />
            &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CommandName="Cancel" CssClass="button-Red" Text="取消" Width="120px" />
            <asp:UpdatePanel ID="upDataEdit" runat="server">
                <ContentTemplate>
                    <table class="TableSetting">
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbCaseNo_Edit" runat="server" CssClass="text-Right-Blue" Text="單號：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="eCaseNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbIsClose_Edit" runat="server" CssClass="text-Right-Blue" Text="處理進度" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:DropDownList ID="ddlIsClose_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlIsClose_Edit_SelectedIndexChanged" Width="90%"></asp:DropDownList>
                                <asp:Label ID="eIsClose_Edit" runat="server" Text='<%# Eval("IsClose") %>' Visible="false" />
                            </td>
                            <td class="ColHeight ColWidth-8Col" colspan="2" />
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="開單日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                <asp:Label ID="lbTitle1_Edit" runat="server" CssClass="titleText-S-Blue" Text="被勸導人資料" Width="98%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbEmpNo_Edit" runat="server" CssClass="text-Right-Blue" Text="員工編號：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:TextBox ID="eEmpNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_Edit_TextChanged" Text='<%# Eval("EmpNo") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbEmpName_Edit" runat="server" CssClass="text-Right-Blue" Text="員工姓名：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="eEmpName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Text='<%# Eval("DepNo") %>' Width="35%" />
                                <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbTitle_Edit" runat="server" CssClass="text-Right-Blue" Text="職稱:" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:TextBox ID="eTitle_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eTitle_Edit_TextChanged" Text='<%# Eval("Title") %>' Width="35%" />
                                <asp:Label ID="eTitle_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title_C") %>' Width="55%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                <asp:Label ID="lbTitle2_Edit" runat="server" CssClass="titleText-S-Blue" Text="勸導事由" Width="98%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbCaseDate_Edit" runat="server" CssClass="text-Right-Blue" Text="日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:TextBox ID="eCaseDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbCaseTime_Edit" runat="server" CssClass="text-Right-Blue" Text="時間：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:TextBox ID="eCaseTime_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseTime") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbPosition_Edit" runat="server" CssClass="text-Right-Blue" Text="地點：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                <asp:TextBox ID="ePosition_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Position") %>' Width="97%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                <asp:Label ID="lbCaseNote_Edit" runat="server" CssClass="text-Right-Blue" Text="事由：" Width="90%" />
                            </td>
                            <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                <asp:TextBox ID="eCaseNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("CaseNote") %>' Height="95%" Width="98%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="輔導人補充說明：" Width="90%" />
                            </td>
                            <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Height="95%" Width="98%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                <asp:Label ID="lbDenialReason_Edit" runat="server" CssClass="text-Right-Blue" Text="被勸導人拒簽原因(可由執行人員依事實填入)：" Width="90%" />
                            </td>
                            <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                <asp:TextBox ID="eDenialReason_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("DenialReason") %>' Height="95%" Width="98%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
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
            <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" OnClick="bbOK_INS_Click" CssClass="button-Black" Text="確定" Width="120px" />
            &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CommandName="Cancel" CssClass="button-Red" Text="取消" Width="120px" />
            <asp:UpdatePanel ID="upDataInsert" runat="server">
                <ContentTemplate>
                    <table class="TableSetting">
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbCaseNo_INS" runat="server" CssClass="text-Right-Blue" Text="單號：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="eCaseNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbIsClose_INS" runat="server" CssClass="text-Right-Blue" Text="處理進度" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:DropDownList ID="ddlIsClose_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlIsClose_INS_SelectedIndexChanged" Width="90%"></asp:DropDownList>
                                <asp:Label ID="eIsClose_INS" runat="server" Text='<%# Eval("IsClose") %>' Visible="false" />
                            </td>
                            <td class="ColHeight ColWidth-8Col" colspan="2" />
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="開單日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                <asp:Label ID="lbTitle1_INS" runat="server" CssClass="titleText-S-Blue" Text="被勸導人資料" Width="98%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbEmpNo_INS" runat="server" CssClass="text-Right-Blue" Text="員工編號：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:TextBox ID="eEmpNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_INS_TextChanged" Text='<%# Eval("EmpNo") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbEmpName_INS" runat="server" CssClass="text-Right-Blue" Text="員工姓名：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="eEmpName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Eval("DepNo") %>' Width="35%" />
                                <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbTitle_INS" runat="server" CssClass="text-Right-Blue" Text="職稱:" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:TextBox ID="eTitle_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eTitle_INS_TextChanged" Text='<%# Eval("Title") %>' Width="35%" />
                                <asp:Label ID="eTitle_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title_C") %>' Width="55%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                                <asp:Label ID="lbTitle2_INS" runat="server" CssClass="titleText-S-Blue" Text="勸導事由" Width="98%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbCaseDate_INS" runat="server" CssClass="text-Right-Blue" Text="日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:TextBox ID="eCaseDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbCaseTime_INS" runat="server" CssClass="text-Right-Blue" Text="時間：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:TextBox ID="eCaseTime_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseTime") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbPosition_INS" runat="server" CssClass="text-Right-Blue" Text="地點：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                <asp:TextBox ID="ePosition_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Position") %>' Width="97%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                <asp:Label ID="lbCaseNote_INS" runat="server" CssClass="text-Right-Blue" Text="事由：" Width="90%" />
                            </td>
                            <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                <asp:TextBox ID="eCaseNote_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("CaseNote") %>' Height="95%" Width="98%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="輔導人補充說明：" Width="90%" />
                            </td>
                            <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Height="95%" Width="98%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                <asp:Label ID="lbDenialReason_INS" runat="server" CssClass="text-Right-Blue" Text="被勸導人拒簽原因(可由執行人員依事實填入)：" Width="90%" />
                            </td>
                            <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                <asp:TextBox ID="eDenialReason_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("DenialReason") %>' Height="95%" Width="98%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col">
                                <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
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
            <asp:Button ID="bbNew_Empty" runat="server" CausesValidation="false" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
        </EmptyDataTemplate>
        <ItemTemplate>
            <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
            &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="Edit" Text="修改" Width="120px" />
            &nbsp;<asp:Button ID="bbPrint_List" runat="server" CausesValidation="False" CssClass="button-Black" OnClick="bbPrint_List_Click" Text="列印" Width="120px" />
            &nbsp;<asp:Button ID="bbDel_List" runat="server" CausesValidation="False" CssClass="button-Red" OnClick="bbDel_List_Click" Text="刪除" Width="120px" />
            <table class="TableSetting">
                <tr>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbCaseNo_List" runat="server" CssClass="text-Right-Blue" Text="單號：" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eCaseNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbIsClose_List" runat="server" CssClass="text-Right-Blue" Text="處理進度" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eIsClose_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("IsClose_C") %>' Width="90%" />
                    </td>
                    <td class="ColHeight ColWidth-8Col" colspan="2">
                        <asp:Label ID="eSourceType_List" runat="server" Text='<%# Eval("SourceType") %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="開單日期：" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                        <asp:Label ID="lbTitle1_List" runat="server" CssClass="titleText-S-Blue" Text="被勸導人資料" Width="98%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbEmpNo_List" runat="server" CssClass="text-Right-Blue" Text="員工編號：" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eEmpNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbEmpName_List" runat="server" CssClass="text-Right-Blue" Text="員工姓名：" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eEmpName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                        <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbTitle_List" runat="server" CssClass="text-Right-Blue" Text="職稱:" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eTitle_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title") %>' Width="35%" />
                        <asp:Label ID="eTitle_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title_C") %>' Width="55%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                        <asp:Label ID="lbTitle2_List" runat="server" CssClass="titleText-S-Blue" Text="勸導事由" Width="98%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbCaseDate_List" runat="server" CssClass="text-Right-Blue" Text="日期：" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eCaseDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbCaseTime_List" runat="server" CssClass="text-Right-Blue" Text="時間：" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eCaseTime_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseTime") %>' Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbPosition_List" runat="server" CssClass="text-Right-Blue" Text="地點：" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                        <asp:Label ID="ePosition_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Position") %>' Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="MultiLine-Low ColBorder ColWidth-8Col">
                        <asp:Label ID="lbCaseNote_List" runat="server" CssClass="text-Right-Blue" Text="事由：" Width="90%" />
                    </td>
                    <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                        <asp:TextBox ID="eCaseNote_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("CaseNote") %>' Height="95%" Width="98%" />
                    </td>
                </tr>
                <tr>
                    <td class="MultiLine-Low ColBorder ColWidth-8Col">
                        <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="輔導人補充說明：" Width="90%" />
                    </td>
                    <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                        <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Height="95%" Width="98%" />
                    </td>
                </tr>
                <tr>
                    <td class="MultiLine-Low ColBorder ColWidth-8Col">
                        <asp:Label ID="lbDenialReason_List" runat="server" CssClass="text-Right-Blue" Text="被勸導人拒簽原因(可由執行人員依事實填入)：" Width="90%" />
                    </td>
                    <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                        <asp:TextBox ID="eDenialReason_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("DenialReason") %>' Height="95%" Width="98%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                        <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                        <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate") %>' Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col">
                        <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                        <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                        <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
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
    <asp:SqlDataSource ID="sdsAdviceReport_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, EmpNo, EmpName, Title, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '員工資料        EMPLOYEE        TITLE') AND (CLASSNO = a.Title)) AS Title_C, CaseDate, CaseTime, Position, CaseNote, Remark, DenialReason, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuManName, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName, SourceType, IsClose, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (CLASSNO = a.IsClose) AND (FKEY = '員工疏失勸導單  AdviceReport    IsClose')) AS IsClose_C FROM AdviceReport AS a WHERE (CaseNo = @CaseNo)"
        DeleteCommand="delete AdviceReport where CaseNo = @CaseNo"
        InsertCommand="INSERT INTO AdviceReport(CaseNo, DepNo, EmpNo, EmpName, Title, CaseDate, CaseTime, Position, CaseNote, Remark, DenialReason, BuDate, BuMan, IsClose) VALUES (@CaseNo, @DepNo, @EmpNo, @EmpName, @Title, @CaseDate, @CaseTime, @Position, @CaseNote, @Remark, @DenialReason, @BuDate, @BuMan, @IsClose)"
        UpdateCommand="UPDATE AdviceReport SET DepNo = @DepNo, EmpNo = @EmpNo, EmpName = @EmpName, Title = @Title, CaseDate = @CaseDate, CaseTime = @CaseTime, Position = @Position, CaseNote = @CaseNote, Remark = @Remark, DenialReason = @DenialReason, ModifyDate = @ModifyDate, ModifyMan = @ModifyMan, IsClose = @IsClose WHERE (CaseNo = @CaseNo)"
        OnDeleted="sdsAdviceReport_Detail_Deleted"
        OnInserted="sdsAdviceReport_Detail_Inserted"
        OnUpdated="sdsAdviceReport_Detail_Updated">
        <DeleteParameters>
            <asp:Parameter Name="CaseNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="EmpNo" />
            <asp:Parameter Name="EmpName" />
            <asp:Parameter Name="Title" />
            <asp:Parameter Name="CaseDate" />
            <asp:Parameter Name="CaseTime" />
            <asp:Parameter Name="Position" />
            <asp:Parameter Name="CaseNote" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="DenialReason" />
            <asp:Parameter Name="BuDate" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="IsClose" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridAdviceReportList" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="EmpNo" />
            <asp:Parameter Name="EmpName" />
            <asp:Parameter Name="Title" />
            <asp:Parameter Name="CaseDate" />
            <asp:Parameter Name="CaseTime" />
            <asp:Parameter Name="Position" />
            <asp:Parameter Name="CaseNote" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="DenialReason" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="IsClose" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:Panel ID="plReport" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCloseReport" runat="server" CssClass="button-Red" OnClick="bbCloseReport_Click" Text="結束預覽" Width="120px" />
        <rsweb:ReportViewer ID="rvPrint" runat="server" BackColor="" ClientIDMode="AutoID" HighlightBackgroundColor=""
            InternalBorderColor="204, 204, 204" InternalBorderStyle="Solid" InternalBorderWidth="1px"
            LinkActiveColor="" LinkActiveHoverColor="" LinkDisabledColor=""
            PrimaryButtonBackgroundColor="" PrimaryButtonForegroundColor="" PrimaryButtonHoverBackgroundColor="" PrimaryButtonHoverForegroundColor=""
            SecondaryButtonBackgroundColor="" SecondaryButtonForegroundColor="" SecondaryButtonHoverBackgroundColor="" SecondaryButtonHoverForegroundColor=""
            SplitterBackColor=""
            ToolbarDividerColor="" ToolbarForegroundColor="" ToolbarForegroundDisabledColor=""
            ToolbarHoverBackgroundColor="" ToolbarHoverForegroundColor=""
            ToolBarItemBorderColor="" ToolBarItemBorderStyle="Solid" ToolBarItemBorderWidth="1px" ToolBarItemHoverBackColor=""
            ToolBarItemPressedBorderColor="51, 102, 153" ToolBarItemPressedBorderStyle="Solid" ToolBarItemPressedBorderWidth="1px"
            ToolBarItemPressedHoverBackColor="153, 187, 226" Height="600px" Width="100%" PageCountMode="Actual">
            <LocalReport ReportPath="Report\AdviceReportP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
