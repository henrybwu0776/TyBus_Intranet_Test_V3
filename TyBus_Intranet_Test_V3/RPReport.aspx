<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="RPReport.aspx.cs" Inherits="TyBus_Intranet_Test_V3.RPReport" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="RPReportForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">員工獎懲呈報單</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbCaseComeFrom_Search" runat="server" CssClass="text-Right-Blue" Text="案件來源：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:DropDownList ID="ddlCaseComeFrom_Search" runat="server" CssClass="text-Left-Black" Width="95%"                        
                        AutoPostBack="True" OnSelectedIndexChanged="ddlCaseComeFrom_Search_SelectedIndexChanged" />
                    <br />
                    <asp:Label ID="eCaseComeFrom_Search" runat="server" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbCaseType_Search" runat="server" CssClass="text-Right-Blue" Text="案件類別：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:DropDownList ID="ddlCaseType_Search" runat="server" CssClass="text-Left-Black" Width="95%"
                        AutoPostBack="True" OnSelectedIndexChanged="ddlCaseType_Search_SelectedIndexChanged" />
                    <br />
                    <asp:Label ID="eCaseType_Search" runat="server" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbAssignDate_Search" runat="server" CssClass="text-Right-Blue" Text="呈報日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eAssignDate_Search_S" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="~" Width="5%" />
                    <asp:TextBox ID="eAssignDate_Search_E" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbCaseDate_Search" runat="server" CssClass="text-Right-Blue" Text="事件日期：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:TextBox ID="eCaseDate_Search_S" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="Label2" runat="server" CssClass="text-Left-Black" Text="~" Width="5%" />
                    <asp:TextBox ID="eCaseDate_Search_E" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Search_TextChanged" Width="95%" />
                    <br />
                    <asp:Label ID="eDepName_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbEmpNo_Search" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:TextBox ID="eEmpNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_Search_TextChanged" Width="95%" />
                    <br />
                    <asp:Label ID="eEmpName_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbCar_ID_Search" runat="server" CssClass="text-Right-Blue" Text="牌照號碼：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:TextBox ID="eCar_ID_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-9Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="搜尋" Width="95%" OnClick="bbSearch_Click" />
                </td>
                <td class="ColHeight ColWidth-9Col">
                    <asp:Button ID="bbExcel_List" runat="server" CssClass="button-Blue" OnClick="bbExcel_List_Click" Text="匯出 Excel" Width="120px" />
                </td>
                <td class="ColHeight ColWidth-9Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="離開" Width="95%" OnClick="bbClose_Click" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
                <td class="ColWidth-9Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">        
        <asp:GridView ID="gridRPReportList" runat="server" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1"
            GridLines="None" AllowPaging="True" AutoGenerateColumns="False" DataKeyNames="CaseNo" DataSourceID="sdsRPReportList" PageSize="5" Width="100%"
            OnPageIndexChanging="gridRPReportList_PageIndexChanging">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="CaseNo" HeaderText="案號" ReadOnly="True" SortExpression="CaseNo" />
                <asp:BoundField DataField="CaseComeFrom_C" HeaderText="案件來源" ReadOnly="True" SortExpression="CaseComeFrom_C" />
                <asp:BoundField DataField="CaseType_C" HeaderText="案件類別" ReadOnly="True" SortExpression="CaseType_C" />
                <asp:BoundField DataField="DepName" HeaderText="所屬單位" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="EmpNo" HeaderText="員工代號" SortExpression="EmpNo" />
                <asp:BoundField DataField="EmpName" HeaderText="姓名" SortExpression="EmpName" />
                <asp:BoundField DataField="Title_C" HeaderText="級職" SortExpression="Title" />
                <asp:BoundField DataField="CaseDate" DataFormatString="{0:D}" HeaderText="事件日期" SortExpression="CaseDate" />
                <asp:BoundField DataField="Car_ID" HeaderText="車號" SortExpression="Car_ID" />
                <asp:BoundField DataField="Position" HeaderText="地點" SortExpression="Position" />
                <asp:BoundField DataField="CaseNote" HeaderText="事件概述" SortExpression="CaseNote" />
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
        <asp:FormView ID="fvRPReportDetail" runat="server" DataKeyNames="CaseNo" DataSourceID="sdsRPReportDetail" Width="100%" OnDataBound="fvRPReportDetail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" OnClick="bbOK_Edit_Click" Text="更新" CssClass="button-Black" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CommandName="Cancel" Text="取消" CssClass="button-Red" Width="120px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseNo_Edit" runat="server" CssClass="text-Right-Blue" Text="案號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseComeFrom_Edit" runat="server" CssClass="text-Right-Blue" Text="案件來源：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:DropDownList ID="ddlCaseComeFrom_Edit" runat="server" CssClass="text-Left-Black" 
                                AutoPostBack="true" OnSelectedIndexChanged="ddlCaseComeFrom_Edit_SelectedIndexChanged" Width="90%"></asp:DropDownList>
                            <asp:Label ID="eCaseComeFrom_Edit" runat="server" Text='<%# Eval("CaseComeFrom") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseType_Edit" runat="server" CssClass="text-Right-Blue" Text="獎懲類別：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:DropDownList ID="ddlCaseType_Edit" runat="server" CssClass="text-Left-Black" Width="95%" 
                                AutoPostBack="True" OnSelectedIndexChanged="ddlCaseType_Edit_SelectedIndexChanged" />
                            <br />
                            <asp:Label ID="eCaseType_Edit" runat="server" Text='<%# Eval("CaseType") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbEmpNo_Edit" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eEmpNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_Edit_TextChanged" Text='<%# Eval("EmpNo") %>' Width="35%" />
                            <asp:Label ID="eEmpName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbTitle_Edit" runat="server" CssClass="text-Right-Blue" Text="級職：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eTitle_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eTitle_Edit_TextChanged" Text='<%# Eval("Title") %>' Width="35%" />
                            <asp:Label ID="eTitle_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title_C") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseDate_Edit" runat="server" CssClass="text-Right-Blue" Text="事件日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbPosition_Edit" runat="server" CssClass="text-Right-Blue" Text="地點：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="ePosition_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Position") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCar_ID_Edit" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCar_ID_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseNote_Edit" runat="server" CssClass="text-Right-Blue" Text="事件概述：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eCaseNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("CaseNote") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAccordingTerms_Edit" runat="server" CssClass="text-Right-Blue" Text="依據條款：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eAccordingTerms_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("AccordingTerms") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="6">
                            <asp:Label ID="lbReview_Edit" runat="server" CssClass="titleText-Blue" Text="獎懲審議" Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eGiveBounds_Edit" runat="server" CssClass="text-Left-Blue" Text="獎金：" Checked='<%# Eval("GiveBounds") %>' />
                            <asp:TextBox ID="eBoundsAmount_Edit" runat="server" CssClass="text-Left-Blue" Text='<%# Eval("BoundsAmount") %>' Width="50%" />
                            <asp:Label ID="lbSplit1_Edit" runat="server" CssClass="text-Left-Blue" Text=" 元 " />
                        </td>
                        <td class="ColHeight ColWidth-6Col" colspan="2" />
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eCommendation_Edit" runat="server" CssClass="text-Left-Blue" Text="嘉獎：" Checked='<%#Eval("Commendation") %>' />
                            <asp:TextBox ID="eCommendationCount_Edit" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("CommendationCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit2_Edit" runat="server" CssClass="text-Left-Blue" Text=" 次 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eMeritCitation_Edit" runat="server" CssClass="text-Left-Blue" Text="記功：" Checked='<%#Eval("MeritCitation") %>' />
                            <asp:TextBox ID="eMeritCitationCount_Edit" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("MeritCitationCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit3_Edit" runat="server" CssClass="text-Left-Blue" Text=" 次 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eMajorMeritCitation_Edit" runat="server" CssClass="text-Left-Blue" Text="大功：" Checked='<%#Eval("MajorMeritCitation") %>' />
                            <asp:TextBox ID="eMajorMeritCitationCOunt_Edit" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("MajorMeritCitationCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit4_Edit" runat="server" CssClass="text-Left-Blue" Text=" 次 " />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eAskExact_Edit" runat="server" CssClass="text-Left-Red" Text="賠償：" Checked='<%#Eval("AskExact") %>' />
                            <asp:TextBox ID="eExactAmount_Edit" runat="server" CssClass="text-Left-Red" Text='<%#Eval("ExactAmount") %>' Width="50%" />
                            <asp:Label ID="lbSplit5_Edit" runat="server" CssClass="text-Left-Red" Text=" 元 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eAdvice_Edit" runat="server" CssClass="text-Left-Red" Text="勸導" Checked='<%#Eval("Advice") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eAdmonition_Edit" runat="server" CssClass="text-Left-Red" Text="口頭警告" Checked='<%#Eval("Admonition") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eReprimand_Edit" runat="server" CssClass="text-Left-Red" Text="申誡：" Checked='<%#Eval("Reprimand") %>' />
                            <asp:TextBox ID="eReprimandCount_Edit" runat="server" CssClass="text-Left-Red" Text='<%#Eval("ReprimandCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit6_Edit" runat="server" CssClass="text-Left-Red" Text=" 次 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eDemerit_Edit" runat="server" CssClass="text-Left-Red" Text="記過：" Checked='<%#Eval("Demerit") %>' />
                            <asp:TextBox ID="eDemeritCount_Edit" runat="server" CssClass="text-Left-Red" Text='<%#Eval("DemeritCount") %>' Width="50%" />
                            <asp:Label ID="lbSPlit7_Edit" runat="server" CssClass="text-Left-Red" Text=" 次 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eMajorDemerit_Edit" runat="server" CssClass="text-Left-Red" Text="大過：" Checked='<%#Eval("MajorDemerit") %>' />
                            <asp:TextBox ID="eMajorDemeritCount_Edit" runat="server" CssClass="text-Left-Red" Text='<%#Eval("MajorDemeritCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit8_Edit" runat="server" CssClass="text-Left-Red" Text=" 次 " />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="ePromotion_Edit" runat="server" CssClass="text-Left-Black" Text="晉升" Checked='<%#Eval("Promotion") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eDemotion_Edit" runat="server" CssClass="text-Left-Black" Text="降調" Checked='<%#Eval("Demotion") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eDismissal_Edit" runat="server" CssClass="text-Left-Black" Text="解職" Checked='<%#Eval("Dismissal") %>' />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="3">
                            <asp:CheckBox ID="eOthers_Edit" runat="server" CssClass="text-Left-Black" Text="其他：" Checked='<%#Eval("Others") %>' />
                            <asp:TextBox ID="eReview_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Review") %>' Height="95%" Width="85%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAssignDate_Edit" runat="server" CssClass="text-Right-Blue" Text="呈報日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eAssignDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAssignDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="呈報單位：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eAssignDepNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDepNo") %>' Width="35%" />
                            <asp:Label ID="eAssignDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAssignMan_Edit" runat="server" CssClass="text-Right-Blue" Text="呈報人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eAssignMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' Width="35%" />
                            <asp:Label ID="eAssignManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCustomServiceNo_Edit" runat="server" CssClass="text-Right-Blue" Text="對應客服單號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCustomServiceNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CustomServiceNo") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
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
                </table></ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" OnClick="bbOK_INS_Click" Text="確定" CssClass="button-Black" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CommandName="Cancel" Text="取消" CssClass="button-Red" Width="120px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseNo_INS" runat="server" CssClass="text-Right-Blue" Text="案號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseComeFrom_INS" runat="server" CssClass="text-Right-Blue" Text="案件來源：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:DropDownList ID="ddlCaseComeFrom_INS" runat="server" CssClass="text-Left-Black" 
                                AutoPostBack="true" OnSelectedIndexChanged="ddlCaseComeFrom_INS_SelectedIndexChanged" Width="90%"></asp:DropDownList>
                            <asp:Label ID="eCaseComeFrom_INS" runat="server" Text='<%# Eval("CaseComeFrom") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseType_INS" runat="server" CssClass="text-Right-Blue" Text="獎懲類別：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:DropDownList ID="ddlCaseType_INS" runat="server" CssClass="text-Left-Black" Width="95%"
                                AutoPostBack="True" OnSelectedIndexChanged="ddlCaseType_INS_SelectedIndexChanged" />
                            <asp:Label ID="eCaseType_INS" runat="server" Text='<%# Eval("CaseType") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbEmpNo_INS" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eEmpNo_INS" runat="server" CssClass="text-Left-Black" OnTextChanged="eEmpNo_INS_TextChanged" AutoPostBack="true" Text='<%# Eval("EmpNo") %>' Width="35%" />
                            <asp:Label ID="eEmpName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" OnTextChanged="eDepNo_INS_TextChanged" AutoPostBack="true" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbTitle_INS" runat="server" CssClass="text-Right-Blue" Text="級職：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eTitle_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eTitle_INS_TextChanged" Text='<%# Eval("Title") %>' Width="35%" />
                            <asp:Label ID="eTitle_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title_C") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseDate_INS" runat="server" CssClass="text-Right-Blue" Text="事件日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbPosition_INS" runat="server" CssClass="text-Right-Blue" Text="地點：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="ePosition_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Position") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCar_ID_INS" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCar_ID_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseNote_INS" runat="server" CssClass="text-Right-Blue" Text="事件概述：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eCaseNote_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("CaseNote") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAccordingTerms_INS" runat="server" CssClass="text-Right-Blue" Text="依據條款：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eAccordingTerms_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("AccordingTerms") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="6">
                            <asp:Label ID="lbReview_INS" runat="server" CssClass="titleText-Blue" Text="獎懲審議" Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eGiveBounds_INS" runat="server" CssClass="text-Left-Blue" Text="獎金：" />
                            <asp:TextBox ID="eBoundsAmount_INS" runat="server" CssClass="text-Left-Blue" Text='<%# Eval("BoundsAmount") %>' Width="50%" />
                            <asp:Label ID="lbSplit1_INS" runat="server" CssClass="text-Left-Blue" Text=" 元 " />
                        </td>
                        <td class="ColHeight ColWidth-6Col" colspan="2" />
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eCommendation_INS" runat="server" CssClass="text-Left-Blue" Text="嘉獎：" />
                            <asp:TextBox ID="eCommendationCount_INS" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("CommendationCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit2_INS" runat="server" CssClass="text-Left-Blue" Text=" 次 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eMeritCitation_INS" runat="server" CssClass="text-Left-Blue" Text="記功：" />
                            <asp:TextBox ID="eMeritCitationCount_INS" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("MeritCitationCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit3_INS" runat="server" CssClass="text-Left-Blue" Text=" 次 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eMajorMeritCitation_INS" runat="server" CssClass="text-Left-Blue" Text="大功：" />
                            <asp:TextBox ID="eMajorMeritCitationCOunt_INS" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("MajorMeritCitationCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit4_INS" runat="server" CssClass="text-Left-Blue" Text=" 次 " />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eAskExact_INS" runat="server" CssClass="text-Left-Red" Text="賠償：" />
                            <asp:TextBox ID="eExactAmount_INS" runat="server" CssClass="text-Left-Red" Text='<%#Eval("ExactAmount") %>' Width="50%" />
                            <asp:Label ID="lbSplit5_INS" runat="server" CssClass="text-Left-Red" Text=" 元 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eAdvice_INS" runat="server" CssClass="text-Left-Red" Text="勸導" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eAdmonition_INS" runat="server" CssClass="text-Left-Red" Text="口頭警告" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eReprimand_INS" runat="server" CssClass="text-Left-Red" Text="申誡：" />
                            <asp:TextBox ID="eReprimandCount_INS" runat="server" CssClass="text-Left-Red" Text='<%#Eval("ReprimandCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit6_INS" runat="server" CssClass="text-Left-Red" Text=" 次 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eDemerit_INS" runat="server" CssClass="text-Left-Red" Text="記過：" />
                            <asp:TextBox ID="eDemeritCount_INS" runat="server" CssClass="text-Left-Red" Text='<%#Eval("DemeritCount") %>' Width="50%" />
                            <asp:Label ID="lbSPlit7_INS" runat="server" CssClass="text-Left-Red" Text=" 次 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eMajorDemerit_INS" runat="server" CssClass="text-Left-Red" Text="大過：" />
                            <asp:TextBox ID="eMajorDemeritCount_INS" runat="server" CssClass="text-Left-Red" Text='<%#Eval("MajorDemeritCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit8_INS" runat="server" CssClass="text-Left-Red" Text=" 次 " />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="ePromotion_INS" runat="server" CssClass="text-Left-Black" Text="晉升" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eDemotion_INS" runat="server" CssClass="text-Left-Black" Text="降調" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eDismissal_INS" runat="server" CssClass="text-Left-Black" Text="解職" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="3">
                            <asp:CheckBox ID="eOthers_INS" runat="server" CssClass="text-Left-Black" Text="其他：" />
                            <asp:TextBox ID="eReview_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Review") %>' Height="95%" Width="85%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAssignDate_INS" runat="server" CssClass="text-Right-Blue" Text="呈報日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eAssignDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAssignDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="呈報單位：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eAssignDepNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDepNo") %>' Width="35%" />
                            <asp:Label ID="eAssignDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAssignMan_INS" runat="server" CssClass="text-Right-Blue" Text="呈報人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eAssignMan_INS" runat="server" CssClass="text-Left-Black" OnTextChanged="eAssignMan_INS_TextChanged" AutoPostBack="true" Text='<%# Eval("AssignMan") %>' Width="35%" />
                            <asp:Label ID="eAssignManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCustomServiceNo_INS" runat="server" CssClass="text-Right-Blue" Text="對應客服單號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCustomServiceNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CustomServiceNo") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
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
                </table></ContentTemplate>
                </asp:UpdatePanel>
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CausesValidation="False" CommandName="New" Text="新增" CssClass="button-Black" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CommandName="New" Text="新增" CssClass="button-Black" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CommandName="Edit" Text="編輯" CssClass="button-Black" Width="120px" />
                &nbsp;<asp:Button ID="bbPrint_List" runat="server" CausesValidation="False" OnClick="bbPrint_List_Click" Text="列印" CssClass="button-Red" Width="120px" />
                &nbsp;<asp:Button ID="bbDEL_List" runat="server" CausesValidation="False" OnClick="bbDEL_List_Click" Text="刪除" CssClass="button-Red" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseNo_List" runat="server" CssClass="text-Right-Blue" Text="案號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseComeFrom_List" runat="server" CssClass="text-Right-Blue" Text="案件來源：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseComeFrom_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseComeFrom_C") %>' Width="95%" />
                            <asp:Label ID="eCaseComeFrom_List" runat="server" Text='<%# Eval("CaseComeFrom") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseType_List" runat="server" CssClass="text-Right-Blue" Text="獎懲類別：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseType_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseType_C") %>' Width="95%" />
                            <asp:Label ID="eCaseType_List" runat="server" Text='<%# Eval("CaseType") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbEmpNo_List" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eEmpNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="35%" />
                            <asp:Label ID="eEmpName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbTitle_List" runat="server" CssClass="text-Right-Blue" Text="級職：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eTitle_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title") %>' Width="35%" />
                            <asp:Label ID="eTitle_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title_C") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseDate_List" runat="server" CssClass="text-Right-Blue" Text="事件日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbPosition_List" runat="server" CssClass="text-Right-Blue" Text="地點：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="ePosition_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Position") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCar_ID_List" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCar_ID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseNote_List" runat="server" CssClass="text-Right-Blue" Text="事件概述：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eCaseNote_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("CaseNote") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAccordingTerms_List" runat="server" CssClass="text-Right-Blue" Text="依據條款：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eAccordingTerms_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("AccordingTerms") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="6">
                            <asp:Label ID="lbReview_List" runat="server" CssClass="titleText-Blue" Text="獎懲審議" Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eGiveBounds_List" runat="server" CssClass="text-Left-Blue" Text="獎金：" Checked='<%# Eval("GiveBounds") %>' Enabled="false" />
                            <asp:Label ID="eBoundsAmount_List" runat="server" CssClass="text-Left-Blue" Text='<%# Eval("BoundsAmount") %>' Width="50%" />
                            <asp:Label ID="lbSplit1_List" runat="server" CssClass="text-Left-Blue" Text=" 元 " />
                        </td>
                        <td class="ColHeight ColWidth-6Col" colspan="2" />
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eCommendation_List" runat="server" CssClass="text-Left-Blue" Text="嘉獎：" Checked='<%#Eval("Commendation") %>' Enabled="false" />
                            <asp:Label ID="eCommendationCount_List" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("CommendationCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit2_List" runat="server" CssClass="text-Left-Blue" Text=" 次 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eMeritCitation_List" runat="server" CssClass="text-Left-Blue" Text="記功：" Checked='<%#Eval("MeritCitation") %>' Enabled="false" />
                            <asp:Label ID="eMeritCitationCount_List" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("MeritCitationCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit3_List" runat="server" CssClass="text-Left-Blue" Text=" 次 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eMajorMeritCitation_List" runat="server" CssClass="text-Left-Blue" Text="大功：" Checked='<%#Eval("MajorMeritCitation") %>' Enabled="false" />
                            <asp:Label ID="eMajorMeritCitationCOunt_List" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("MajorMeritCitationCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit4_List" runat="server" CssClass="text-Left-Blue" Text=" 次 " />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eAskExact_List" runat="server" CssClass="text-Left-Red" Text="賠償：" Checked='<%#Eval("AskExact") %>' Enabled="false" />
                            <asp:Label ID="eExactAmount" runat="server" CssClass="text-Left-Red" Text='<%#Eval("ExactAmount") %>' Width="50%" />
                            <asp:Label ID="lbSplit5_List" runat="server" CssClass="text-Left-Red" Text=" 元 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eAdvice_List" runat="server" CssClass="text-Left-Red" Text="勸導" Checked='<%#Eval("Advice") %>' Enabled="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eAdmonition_List" runat="server" CssClass="text-Left-Red" Text="口頭警告" Checked='<%#Eval("Admonition") %>' Enabled="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eReprimand_List" runat="server" CssClass="text-Left-Red" Text="申誡：" Checked='<%#Eval("Reprimand") %>' Enabled="false" />
                            <asp:Label ID="eReprimandCount_List" runat="server" CssClass="text-Left-Red" Text='<%#Eval("ReprimandCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit6_List" runat="server" CssClass="text-Left-Red" Text=" 次 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eDemerit_List" runat="server" CssClass="text-Left-Red" Text="記過：" Checked='<%#Eval("Demerit") %>' Enabled="false" />
                            <asp:Label ID="eDemeritCount_List" runat="server" CssClass="text-Left-Red" Text='<%#Eval("DemeritCount") %>' Width="50%" />
                            <asp:Label ID="lbSPlit7_List" runat="server" CssClass="text-Left-Red" Text=" 次 " />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eMajorDemerit_List" runat="server" CssClass="text-Left-Red" Text="大過：" Checked='<%#Eval("MajorDemerit") %>' Enabled="false" />
                            <asp:Label ID="eMajorDemeritCount_List" runat="server" CssClass="text-Left-Red" Text='<%#Eval("MajorDemeritCount") %>' Width="50%" />
                            <asp:Label ID="lbSplit8_List" runat="server" CssClass="text-Left-Red" Text=" 次 " />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="ePromotion_List" runat="server" CssClass="text-Left-Black" Text="晉升" Checked='<%#Eval("Promotion") %>' Enabled="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eDemotion_List" runat="server" CssClass="text-Left-Black" Text="降調" Checked='<%#Eval("Demotion") %>' Enabled="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:CheckBox ID="eDismissal_List" runat="server" CssClass="text-Left-Black" Text="解職" Checked='<%#Eval("Dismissal") %>' Enabled="false" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="3">
                            <asp:CheckBox ID="eOthers_List" runat="server" CssClass="text-Left-Black" Text="其他：" Checked='<%#Eval("Others") %>' Enabled="false" />
                            <asp:TextBox ID="eReview_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Review") %>' Height="95%" Width="85%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAssignDate_List" runat="server" CssClass="text-Right-Blue" Text="呈報日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eAssignDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAssignDepNo_List" runat="server" CssClass="text-Right-Blue" Text="呈報單位：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eAssignDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDepNo") %>' Width="35%" />
                            <asp:Label ID="eAssignDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAssignMan_List" runat="server" CssClass="text-Right-Blue" Text="呈報人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eAssignMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignMan") %>' Width="35%" />
                            <asp:Label ID="eAssignManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignManName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCustomServiceNo_List" runat="server" CssClass="text-Right-Blue" Text="對應客服單號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCustomServiceNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CustomServiceNo") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
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
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
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
            <LocalReport ReportPath="Report\RP_ReportP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsRPReportList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = r.CaseComeFrom) AND (FKEY = '員工獎懲呈報單  RPReport        CaseComeFrom')) AS CaseComeFrom_C, (SELECT CLASSTXT FROM DBDICB AS DBDICB_2 WHERE (FKEY = '員工獎懲呈報單  RPReport        CaseType') AND (CLASSNO = r.CaseType)) AS CaseType_C, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = r.DepNo)) AS DepName, EmpNo, EmpName, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (CLASSNO = r.Title) AND (FKEY = '人事資料檔      EMPLOYEE        TITLE')) AS Title_C, CaseDate, Car_ID, Position, CaseNote FROM RP_Report AS r WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsRPReportDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, CaseComeFrom, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = r.CaseComeFrom) AND (FKEY = '員工獎懲呈報單  RPReport        CaseComeFrom')) AS CaseComeFrom_C, CaseType, (SELECT CLASSTXT FROM DBDICB AS DBDICB_2 WHERE (FKEY = '員工獎懲呈報單  RPReport        CaseType') AND (CLASSNO = r.CaseType)) AS CaseType_C, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = r.DepNo)) AS DepName, EmpNo, EmpName, Title, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (CLASSNO = r.Title) AND (FKEY = '人事資料檔      EMPLOYEE        TITLE')) AS Title_C, CaseDate, Car_ID, Position, CaseNote, AccordingTerms, GiveBounds, BoundsAmount, AskExact, ExactAmount, Demotion, Dismissal, Promotion, Advice, Admonition, Reprimand, ReprimandCount, Demerit, DemeritCount, MajorDemerit, MajorDemeritCount, Commendation, CommendationCount, MeritCitation, MeritCitationCount, MajorMeritCitation, MajorMeritCitationCount, Others, Review, Remark, AssignDate, AssignDepNo, (SELECT NAME FROM DEPARTMENT AS DEPARTMENT_1 WHERE (DEPNO = r.AssignDepNo)) AS AssignDepName, AssignMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = r.AssignMan)) AS AssignManName, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = r.BuMan)) AS BuManName, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = r.ModifyMan)) AS ModifyManName, CustomServiceNo FROM RP_Report AS r WHERE (CaseNo = @CaseNo)"
        DeleteCommand="DELETE FROM RP_Report WHERE (CaseNo = @CaseNo)"
        InsertCommand="INSERT INTO RP_Report(CaseNo, CaseComeFrom, CaseType, DepNo, EmpNo, EmpName, Title, CaseDate, Car_ID, Position, CaseNote, AccordingTerms, Review, Remark, AssignDate, AssignDepNo, AssignMan, BuDate, BuMan, GiveBounds, BoundsAmount, AskExact, ExactAmount, Demotion, Dismissal, Promotion, Advice, Admonition, Reprimand, ReprimandCount, Demerit, DemeritCount, MajorDemerit, MajorDemeritCount, Commendation, CommendationCount, MeritCitation, MeritCitationCount, MajorMeritCitation, MajorMeritCitationCount, Others) VALUES (@CaseNo, @CaseComeFrom, @CaseType, @DepNo, @EmpNo, @EmpName, @Title, @CaseDate, @Car_ID, @Position, @CaseNote, @AccordingTerms, @Review, @Remark, @AssignDate, @AssignDepNo, @AssignMan, @BuDate, @BuMan, @GiveBounds, @BoundsAmount, @AskExact, @ExactAmount, @Demotion, @Dismissal, @Promotion, @Advice, @Admonition, @Reprimand, @ReprimandCount, @Demerit, @DemeritCount, @MajorDemerit, @MajorDemeritCount, @Commendation, @CommendationCount, @MeritCitation, @MeritCitationCount, @MajorMeritCitation, @MajorMeritCitationCount, @Others)"
        UpdateCommand="UPDATE RP_Report SET CaseComeFrom = @CaseComeFrom, CaseType = @CaseType, DepNo = @DepNo, EmpNo = @EmpNo, EmpName = @EmpName, Title = @Title, CaseDate = @CaseDate, Car_ID = @Car_ID, Position = @Position, CaseNote = @CaseNote, AccordingTerms = @AccordingTerms, Review = @Review, Remark = @Remark, ModifyDate = @ModifyDate, ModifyMan = @ModifyMan, GiveBounds = @GiveBounds, BoundsAmount = @BoundsAmount, AskExact = @AskExact, ExactAmount = @ExactAmount, Demotion = @Demotion, Dismissal = @Dismissal, Promotion = @Promotion, Advice = @Advice, Admonition = @Admonition, Reprimand = @Reprimand, ReprimandCount = @ReprimandCount, Demerit = @Demerit, DemeritCount = @DemeritCount, MajorDemerit = @MajorDemerit, MajorDemeritCount = @MajorDemeritCount, Commendation = @Commendation, CommendationCount = @CommendationCount, MeritCitation = @MeritCitation, MeritCitationCount = @MeritCitationCount, MajorMeritCitation = @MajorMeritCitation, MajorMeritCitationCount = @MajorMeritCitationCount, Others = @Others WHERE (CaseNo = @CaseNo)">
        <DeleteParameters>
            <asp:Parameter Name="CaseNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="CaseComeFrom" />
            <asp:Parameter Name="CaseType" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="EmpNo" />
            <asp:Parameter Name="EmpName" />
            <asp:Parameter Name="Title" />
            <asp:Parameter Name="CaseDate" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="Position" />
            <asp:Parameter Name="CaseNote" />
            <asp:Parameter Name="AccordingTerms" />
            <asp:Parameter Name="Review" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="AssignDate" />
            <asp:Parameter Name="AssignDepNo" />
            <asp:Parameter Name="AssignMan" />
            <asp:Parameter Name="BuDate" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="GiveBounds" />
            <asp:Parameter Name="BoundsAmount" />
            <asp:Parameter Name="AskExact" />
            <asp:Parameter Name="ExactAmount" />
            <asp:Parameter Name="Demotion" />
            <asp:Parameter Name="Dismissal" />
            <asp:Parameter Name="Promotion" />
            <asp:Parameter Name="Advice" />
            <asp:Parameter Name="Admonition" />
            <asp:Parameter Name="Reprimand" />
            <asp:Parameter Name="ReprimandCount" />
            <asp:Parameter Name="Demerit" />
            <asp:Parameter Name="DemeritCount" />
            <asp:Parameter Name="MajorDemerit" />
            <asp:Parameter Name="MajorDemeritCount" />
            <asp:Parameter Name="Commendation" />
            <asp:Parameter Name="CommendationCount" />
            <asp:Parameter Name="MeritCitation" />
            <asp:Parameter Name="MeritCitationCount" />
            <asp:Parameter Name="MajorMeritCitation" />
            <asp:Parameter Name="MajorMeritCitationCount" />
            <asp:Parameter Name="Others" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridRPReportList" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="CaseComeFrom" />
            <asp:Parameter Name="CaseType" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="EmpNo" />
            <asp:Parameter Name="EmpName" />
            <asp:Parameter Name="Title" />
            <asp:Parameter Name="CaseDate" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="Position" />
            <asp:Parameter Name="CaseNote" />
            <asp:Parameter Name="AccordingTerms" />
            <asp:Parameter Name="Review" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="GiveBounds" />
            <asp:Parameter Name="BoundsAmount" />
            <asp:Parameter Name="AskExact" />
            <asp:Parameter Name="ExactAmount" />
            <asp:Parameter Name="Demotion" />
            <asp:Parameter Name="Dismissal" />
            <asp:Parameter Name="Promotion" />
            <asp:Parameter Name="Advice" />
            <asp:Parameter Name="Admonition" />
            <asp:Parameter Name="Reprimand" />
            <asp:Parameter Name="ReprimandCount" />
            <asp:Parameter Name="Demerit" />
            <asp:Parameter Name="DemeritCount" />
            <asp:Parameter Name="MajorDemerit" />
            <asp:Parameter Name="MajorDemeritCount" />
            <asp:Parameter Name="Commendation" />
            <asp:Parameter Name="CommendationCount" />
            <asp:Parameter Name="MeritCitation" />
            <asp:Parameter Name="MeritCitationCount" />
            <asp:Parameter Name="MajorMeritCitation" />
            <asp:Parameter Name="MajorMeritCitationCount" />
            <asp:Parameter Name="Others" />
            <asp:Parameter Name="CaseNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
    </asp:Content>
