<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DriverRiskLevel.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DriverRiskLevel" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="DriverRiskLevelForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">員工風險評估等級</a>
    </div>
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="現職單位：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eDepName_Search" runat="server" CssClass="text-Left-Black" Width="60%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbEmpNo_Search" runat="server" CssClass="text-Right-Blue" Text="員工：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eEmpNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eEmpName_Search" runat="server" CssClass="text-Left-Black" Width="60%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:CheckBox ID="cbIsIgnoreLevel2_Search" runat="server" CssClass="text-Left-Black" Text="評級 2 不列入報表" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" rowspan="2">
                    <asp:Label ID="lbRiskDate_Search" runat="server" CssClass="text-Right-Blue" Text="評鑑年度：" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" rowspan="2" colspan="4">
                    <asp:CheckBox ID="cbGetFullYear_Search" runat="server" CssClass="text-Left-Red" Text="取得全年度資料" OnCheckedChanged="cbGetFullYear_Search_CheckedChanged" AutoPostBack="True" />
                    <br />
                    <asp:TextBox ID="eRiskYearS_Search" runat="server" CssClass="text-Left-Black" Width="25%" />
                    <asp:DropDownList ID="eRiskMonthS_Search" runat="server" CssClass="text-Left-Black" Width="15%">
                        <asp:ListItem Text="一月" Value="1" />
                        <asp:ListItem Text="二月" Value="2" />
                        <asp:ListItem Text="三月" Value="3" />
                        <asp:ListItem Text="四月" Value="4" />
                        <asp:ListItem Text="五月" Value="5" />
                        <asp:ListItem Text="六月" Value="6" />
                        <asp:ListItem Text="七月" Value="7" />
                        <asp:ListItem Text="八月" Value="8" />
                        <asp:ListItem Text="九月" Value="9" />
                        <asp:ListItem Text="十月" Value="10" />
                        <asp:ListItem Text="十一月" Value="11" />
                        <asp:ListItem Text="十二月" Value="12" />
                    </asp:DropDownList>
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eRiskYearE_Search" runat="server" CssClass="text-Left-Black" Width="25%" />
                    <asp:DropDownList ID="eRiskMonthE_Search" runat="server" CssClass="text-Left-Black" Width="15%">
                        <asp:ListItem Text="一月" Value="1" />
                        <asp:ListItem Text="二月" Value="2" />
                        <asp:ListItem Text="三月" Value="3" />
                        <asp:ListItem Text="四月" Value="4" />
                        <asp:ListItem Text="五月" Value="5" />
                        <asp:ListItem Text="六月" Value="6" />
                        <asp:ListItem Text="七月" Value="7" />
                        <asp:ListItem Text="八月" Value="8" />
                        <asp:ListItem Text="九月" Value="9" />
                        <asp:ListItem Text="十月" Value="10" />
                        <asp:ListItem Text="十一月" Value="11" />
                        <asp:ListItem Text="十二月" Value="12" />
                    </asp:DropDownList>
                </td>
                <td class="ColHeight ColWidth-8Col" colspan="2">
                    <asp:FileUpload ID="fuExcel" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbImportExcel" runat="server" CssClass="button-Red" OnClick="bbImportExcel_Click" Text="從 EXCEL 匯入" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" OnClick="bbSearch_Click" Text="查詢" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExportExcel" runat="server" CssClass="button-Blue" OnClick="bbExportExcel_Click" Text="匯出 EXCEL" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" OnClick="bbClose_Click" Text="結束" Width="95%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbGetNameList" runat="server" CssClass="button-Black" OnClick="bbGetNameList_Click" Text="批次作業" Width="95%" Visible="False" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbPrint" runat="server" CssClass="button-Black" OnClick="bbPrint_Click" Text="列印報表" Width="95%" Visible="False" />
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
        <asp:GridView ID="gridDriverRiskLevelList" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataSourceID="sdsDriverNameList" ForeColor="#333333" GridLines="None" PageSize="5" Width="100%" OnPageIndexChanging="gridDriverRiskLevelList_PageIndexChanging" OnSelectedIndexChanged="gridDriverRiskLevelList_SelectedIndexChanged">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="IDCardNo" HeaderText="身分證字號" SortExpression="IDCardNo" />
                <asp:BoundField DataField="EMPNO" HeaderText="員工工號" SortExpression="EMPNO" />
                <asp:BoundField DataField="EmpName" HeaderText="員工姓名" SortExpression="EmpName" />
                <asp:BoundField DataField="DEPNO" HeaderText="單位代碼" SortExpression="DEPNO" />
                <asp:BoundField DataField="DepName" HeaderText="所屬單位" ReadOnly="True" SortExpression="DepName" />
            </Columns>
            <EditRowStyle BackColor="#7C6F57" />
            <FooterStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#666666" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#E3EAEB" />
            <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#F8FAFA" />
            <SortedAscendingHeaderStyle BackColor="#246B61" />
            <SortedDescendingCellStyle BackColor="#D4DFE1" />
            <SortedDescendingHeaderStyle BackColor="#15524A" />
        </asp:GridView>
        <asp:SqlDataSource ID="sdsDriverNameList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT d.IDCardNo, e.EMPNO, e.NAME AS EmpName, e.DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DEPNO)) AS DepName FROM DriverRiskLevel AS d LEFT OUTER JOIN EMPLOYEE AS e ON e.IDCARDNO = d.IDCardNo WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    </asp:Panel>
    <asp:Panel ID="plShowData_Detail" runat="server" CssClass="ShowPanel-Detail_B">
        <asp:GridView ID="gridDriverRiskLevelDetail" runat="server" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" AutoGenerateColumns="False"
            DataKeyNames="IndexNo" DataSourceID="sdsDriverRiskLevelList" OnPageIndexChanging="gridDriverRiskLevelDetail_PageIndexChanging" OnSelectedIndexChanged="gridDriverRiskLevelDetail_SelectedIndexChanged">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="IndexNo" HeaderText="序號" SortExpression="IndexNo" ReadOnly="True" />
                <asp:BoundField DataField="CalYear" HeaderText="CalYear" SortExpression="CalYear" Visible="False" />
                <asp:BoundField DataField="CalYM" HeaderText="評鑑年月" SortExpression="CalYM" />
                <asp:BoundField DataField="IDCardNo" HeaderText="身分證字號" SortExpression="IDCardNo" Visible="False" />
                <asp:BoundField DataField="EmpNo" HeaderText="工號" ReadOnly="True" SortExpression="EmpNo" />
                <asp:BoundField DataField="EmpName" HeaderText="姓名" ReadOnly="True" SortExpression="EmpName" />
                <asp:BoundField DataField="DepNo" HeaderText="單位代碼" ReadOnly="True" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="單位" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="RiskDate" DataFormatString="{0:D}" HeaderText="評鑑日期" SortExpression="RiskDate" />
                <asp:BoundField DataField="DoctorLevel" HeaderText="職醫評等" SortExpression="DoctorLevel" />
                <asp:BoundField DataField="CompanyLevel" HeaderText="公司評等" SortExpression="CompanyLevel" />
                <asp:BoundField DataField="WorkHours" HeaderText="月工時" SortExpression="WorkHours" />
                <asp:BoundField DataField="WorkDays" HeaderText="天數" SortExpression="WorkDays" />
                <asp:BoundField DataField="Remark" HeaderText="Remark" SortExpression="Remark" Visible="False" />
            </Columns>
            <EditRowStyle BackColor="#2461BF" />
            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#EFF3FB" />
            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#F5F7FB" />
            <SortedAscendingHeaderStyle BackColor="#6D95E1" />
            <SortedDescendingCellStyle BackColor="#E9EBEF" />
            <SortedDescendingHeaderStyle BackColor="#4870BE" />
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="plShowData_Detail2" runat="server" CssClass="ShowPanel-Detail_C">
        <asp:FormView ID="fvDriverRiskLevelDetail" runat="server" DataKeyNames="IndexNo" DataSourceID="sdsDriverRiskLevelDetail" Width="100%" OnDataBound="fvDriverRiskLevelDetail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CssClass="button-Blue" CausesValidation="True" OnClick="bbOK_Edit_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbCalYear_Edit" runat="server" CssClass="text-Right-Blue" Text="歸屬年度：" Width="95%" />
                                    <asp:Label ID="eIndexNo_Edit" runat="server" Text='<%# Eval("IndexNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eCalYear_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CalYear") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbCalYM_Edit" runat="server" CssClass="text-Right-Blue" Text="評鑑年月：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eCalYM_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CalYM") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbRiskDate_Edit" runat="server" CssClass="text-Right-Blue" Text="評鑑日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:TextBox ID="eRiskDate_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" Text='<%# Eval("RiskDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="95%" />
                                    <asp:Label ID="eDepNo_Edit" runat="server" Text='<%# Eval("DepNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbEmpNo_Edit" runat="server" CssClass="text-Right-Blue" Text="受評人員：" Width="95%" />
                                    <asp:Label ID="eIDCardNo_Edit" runat="server" Text='<%# Eval("IDCardNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col" colspan="3">
                                    <asp:Label ID="eEmpNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="35%" />
                                    <asp:Label ID="eEmpName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbDoctorLevel_Edit" runat="server" CssClass="text-Right-Blue" Text="職醫評等：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:TextBox ID="eDoctorLevel_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DoctorLevel") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbCompanyLevel_Edit" runat="server" CssClass="text-Right-Blue" Text="公司評等：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:TextBox ID="eCompanyLevel_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompanyLevel") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbWorkHours_Edit" runat="server" CssClass="text-Right-Blue" Text="月工時：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:TextBox ID="eWorkHours_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("WorkHours") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbWorkDays_Edit" runat="server" CssClass="text-Right-Blue" Text="出勤天數：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:TextBox ID="eWorkDays_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("WorkDays") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_Low ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-4Col" colspan="3">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="95%" />
                                    <asp:Label ID="eBuMan_Edit" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                                    <asp:Label ID="eModifyMan_Edit" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-4Col" />
                                <td class="ColWidth-4Col" />
                                <td class="ColWidth-4Col" />
                                <td class="ColWidth-4Col" />
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CssClass="button-Blue" CausesValidation="True" OnClick="bbOK_INS_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbCalYear_INS" runat="server" CssClass="text-Right-Blue" Text="歸屬年度：" Width="95%" />
                                    <asp:Label ID="eIndexNo_INS" runat="server" Text='<%# Eval("IndexNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eCalYear_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CalYear") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbCalYM_INS" runat="server" CssClass="text-Right-Blue" Text="評鑑年月：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eCalYM_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CalYM") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbRiskDate_INS" runat="server" CssClass="text-Right-Blue" Text="評鑑日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:TextBox ID="eRiskDate_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" Text='<%# Eval("RiskDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="95%" />
                                    <asp:Label ID="eDepNo_INS" runat="server" Text='<%# Eval("DepNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbEmpNo_INS" runat="server" CssClass="text-Right-Blue" Text="受評人員：" Width="95%" />
                                    <asp:Label ID="eIDCardNo_INS" runat="server" Text='<%# Eval("IDCardNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col" colspan="3">
                                    <asp:Label ID="eEmpNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="35%" />
                                    <asp:Label ID="eEmpName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbDoctorLevel_INS" runat="server" CssClass="text-Right-Blue" Text="職醫評等：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:TextBox ID="eDoctorLevel_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DoctorLevel") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbCompanyLevel_INS" runat="server" CssClass="text-Right-Blue" Text="公司評等：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:TextBox ID="eCompanyLevel_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompanyLevel") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbWorkHours_INS" runat="server" CssClass="text-Right-Blue" Text="月工時：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:TextBox ID="eWorkHours_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("WorkHours") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbWorkDays_INS" runat="server" CssClass="text-Right-Blue" Text="出勤天數：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:TextBox ID="eWorkDays_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("WorkDays") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_Low ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-4Col" colspan="3">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="95%" />
                                    <asp:Label ID="eBuMan_INS" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                                    <asp:Label ID="eModifyMan_INS" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-4Col" />
                                <td class="ColWidth-4Col" />
                                <td class="ColWidth-4Col" />
                                <td class="ColWidth-4Col" />
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CssClass="button-Black" CausesValidation="false" CommandName="New" Text="新增" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_List" runat="server" CssClass="button-Black" CausesValidation="False" CommandName="New" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CssClass="button-Blue" CausesValidation="False" CommandName="Edit" Text="編輯" Width="120px" />
                &nbsp;<asp:Button ID="bbDel_List" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Delete" Text="刪除" Width="120px" />
                <asp:UpdatePanel runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbCalYear_List" runat="server" CssClass="text-Right-Blue" Text="歸屬年度：" Width="95%" />
                                    <asp:Label ID="eIndexNo_List" runat="server" Text='<%# Eval("IndexNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eCalYear_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CalYear") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbCalYM_List" runat="server" CssClass="text-Right-Blue" Text="評鑑年月：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eCalYM_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CalYM") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbRiskDate_List" runat="server" CssClass="text-Right-Blue" Text="評鑑日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eRiskDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RiskDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="95%" />
                                    <asp:Label ID="eDepNo_List" runat="server" Text='<%# Eval("DepNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbEmpNo_List" runat="server" CssClass="text-Right-Blue" Text="受評人員：" Width="95%" />
                                    <asp:Label ID="eIDCardNo_List" runat="server" Text='<%# Eval("IDCardNo") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col" colspan="3">
                                    <asp:Label ID="eEmpNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="35%" />
                                    <asp:Label ID="eEmpName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="60%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbDoctorLevel_List" runat="server" CssClass="text-Right-Blue" Text="職醫評等：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eDoctorLevel_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DoctorLevel") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbCompanyLevel_List" runat="server" CssClass="text-Right-Blue" Text="公司評等：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eCompanyLevel_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompanyLevel") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbWorkHours_List" runat="server" CssClass="text-Right-Blue" Text="月工時：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eWorkHours_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("WorkHours") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbWorkDays_List" runat="server" CssClass="text-Right-Blue" Text="出勤天數：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eWorkDays_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("WorkDays") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_Low ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-4Col" colspan="3">
                                    <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Width="97%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="95%" />
                                    <asp:Label ID="eBuMan_List" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                                    <asp:Label ID="eModifyMan_List" runat="server" Text='<%# Eval("ModifyMan") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-4Col">
                                    <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-4Col" />
                                <td class="ColWidth-4Col" />
                                <td class="ColWidth-4Col" />
                                <td class="ColWidth-4Col" />
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCLoseReport" runat="server" CssClass="button-Red" Text="關閉報表" OnClick="bbCLoseReport_Click" Width="120px" />
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
            <LocalReport ReportPath="Report\DriverRiskLevelP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsDriverRiskLevelList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT IndexNo, CalYear, CalYM, IDCardNo, (SELECT TOP (1) EMPNO FROM EMPLOYEE WHERE (IDCARDNO = d.IDCardNo) ORDER BY ASSUMEDAY DESC) AS EmpNo, (SELECT TOP (1) NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (IDCARDNO = d.IDCardNo) ORDER BY ASSUMEDAY DESC) AS EmpName, (SELECT TOP (1) DEPNO FROM EMPLOYEE AS EMPLOYEE_2 WHERE (IDCARDNO = d.IDCardNo) ORDER BY ASSUMEDAY DESC) AS DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = (SELECT TOP (1) DEPNO FROM EMPLOYEE AS EMPLOYEE_1 WHERE (IDCARDNO = d.IDCardNo) ORDER BY ASSUMEDAY DESC))) AS DepName, RiskDate, DoctorLevel, CompanyLevel, WorkHours, WorkDays, Remark FROM DriverRiskLevel AS d WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsDriverRiskLevelDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" 
        SelectCommand="SELECT IndexNo, CalYear, CalYM, IDCardNo, (SELECT TOP (1) EMPNO FROM EMPLOYEE WHERE (IDCARDNO = d.IDCardNo) ORDER BY ASSUMEDAY DESC) AS EmpNo, (SELECT TOP (1) NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (IDCARDNO = d.IDCardNo) ORDER BY ASSUMEDAY DESC) AS EmpName, (SELECT TOP (1) DEPNO FROM EMPLOYEE AS EMPLOYEE_2 WHERE (IDCARDNO = d.IDCardNo) ORDER BY ASSUMEDAY DESC) AS DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = (SELECT TOP (1) DEPNO FROM EMPLOYEE AS EMPLOYEE_1 WHERE (IDCARDNO = d.IDCardNo) ORDER BY ASSUMEDAY DESC))) AS DepName, RiskDate, DoctorLevel, CompanyLevel, WorkHours, WorkDays, Remark, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_5 WHERE (EMPNO = d.BuMan)) AS BuManName, BuDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_4 WHERE (EMPNO = d.ModifyMan)) AS ModifyManName, ModifyDate FROM DriverRiskLevel AS d WHERE (ISNULL(IndexNo, '') = '')" 
        DeleteCommand="DELETE FROM DriverRiskLevel WHERE (IndexNo = @IndexNo)" 
        InsertCommand="INSERT INTO DriverRiskLevel(IndexNo, CalYM, IDCardNo, EmpNo, RiskDate01, DoctorLevel01, CompanyLevel01, WorkHours01, WorkDays01, RiskDate02, DoctorLevel02, CompanyLevel02, WorkHours02, WorkDays02, RiskDate03, DoctorLevel03, CompanyLevel03, WorkHours03, WorkDays03, RiskDate04, DoctorLevel04, CompanyLevel04, WorkHours04, WorkDays04, RiskDate05, DoctorLevel05, CompanyLevel05, WorkHours05, WorkDays05, RiskDate06, DoctorLevel06, CompanyLevel06, WorkHours06, WorkDays06, RiskDate07, DoctorLevel07, CompanyLevel07, WorkHours07, WorkDays07, RiskDate08, DoctorLevel08, CompanyLevel08, WorkHours08, WorkDays08, RiskDate09, DoctorLevel09, CompanyLevel09, WorkHours09, WorkDays09, RiskDate10, DoctorLevel10, CompanyLevel10, WorkHours10, WorkDays10, RiskDate11, DoctorLevel11, CompanyLevel11, WorkHours11, WorkDays11, RiskDate12, DoctorLevel12, CompanyLevel12, WorkHours12, WorkDays12, Remark, BuMan, BuDate) VALUES (@IndexNo, @CalYM, @IDCardNo, @EmpNo, @RiskDate01, @DoctorLevel01, @CompanyLevel01, @WorkHours01, @WorkDays01, @RiskDate02, @DoctorLevel02, @CompanyLevel02, @WorkHours02, @WorkDays02, @RiskDate03, @DoctorLevel03, @CompanyLevel03, @WorkHours03, @WorkDays03, @RiskDate04, @DoctorLevel04, @CompanyLevel04, @WorkHours04, @WorkDays04, @RiskDate05, @DoctorLevel05, @CompanyLevel05, @WorkHours05, @WorkDays05, @RiskDate06, @DoctorLevel06, @CompanyLevel06, @WorkHours06, @WorkDays06, @RiskDate07, @DoctorLevel07, @CompanyLevel07, @WorkHours07, @WorkDays07, @RiskDate08, @DoctorLevel08, @CompanyLevel08, @WorkHours08, @WorkDays08, @RiskDate09, @DoctorLevel09, @CompanyLevel09, @WorkHours09, @WorkDays09, @RiskDate10, @DoctorLevel10, @CompanyLevel10, @WorkHours10, @WorkDays10, @RiskDate11, @DoctorLevel11, @CompanyLevel11, @WorkHours11, @WorkDays11, @RiskDate12, @DoctorLevel12, @CompanyLevel12, @WorkHours12, @WorkDays12, @Remark, @BuMan, @BuDate)" 
        UpdateCommand="UPDATE DriverRiskLevel SET RiskDate01 = @RiskDate01, DoctorLevel01 = @DoctorLevel01, CompanyLevel01 = @CompanyLevel01, WorkHours01 = @WorkHours01, WorkDays01 = @WorkDays01, RiskDate02 = @RiskDate02, DoctorLevel02 = @DoctorLevel02, CompanyLevel02 = @CompanyLevel02, WorkHours02 = @WorkHours02, WorkDays02 = @WorkDays02, RiskDate03 = @RiskDate03, DoctorLevel03 = @DoctorLevel03, CompanyLevel03 = @CompanyLevel03, WorkHours03 = @WorkHours03, WorkDays03 = @WorkDays03, RiskDate04 = @RiskDate04, DoctorLevel04 = @DoctorLevel04, CompanyLevel04 = @CompanyLevel04, WorkHours04 = @WorkHours04, WorkDays04 = @WorkDays04, RiskDate05 = @RiskDate05, DoctorLevel05 = @DoctorLevel05, CompanyLevel05 = @CompanyLevel05, WorkHours05 = @WorkHours05, WorkDays05 = @WorkDays05, RiskDate06 = @RiskDate06, DoctorLevel06 = @DoctorLevel06, CompanyLevel06 = @CompanyLevel06, WorkHours06 = @WorkHours06, WorkDays06 = @WorkDays06, RiskDate07 = @RiskDate07, DoctorLevel07 = @DoctorLevel07, CompanyLevel07 = @CompanyLevel07, WorkHours07 = @WorkHours07, WorkDays07 = @WorkDays07, RiskDate08 = @RiskDate08, DoctorLevel08 = @DoctorLevel08, CompanyLevel08 = @CompanyLevel08, WorkHours08 = @WorkHours08, WorkDays08 = @WorkDays08, RiskDate09 = @RiskDate09, DoctorLevel09 = @DoctorLevel09, CompanyLevel09 = @CompanyLevel09, WorkHours09 = @WorkHours09, WorkDays09 = @WorkDays09, RiskDate10 = @RiskDate10, DoctorLevel10 = @DoctorLevel10, CompanyLevel10 = @CompanyLevel10, WorkHours10 = @WorkHours10, WorkDays10 = @WorkDays10, RiskDate11 = @RiskDate11, DoctorLevel11 = @DoctorLevel11, CompanyLevel11 = @CompanyLevel11, WorkHours11 = @WorkHours11, WorkDays11 = @WorkDays11, RiskDate12 = @RiskDate12, DoctorLevel12 = @DoctorLevel12, CompanyLevel12 = @CompanyLevel12, WorkHours12 = @WorkHours12, WorkDays12 = @WorkDays12, Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate WHERE (IndexNo = @IndexNo)">
        <DeleteParameters>
            <asp:Parameter Name="IndexNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="IndexNo" />
            <asp:Parameter Name="CalYM" />
            <asp:Parameter Name="IDCardNo" />
            <asp:Parameter Name="EmpNo" />
            <asp:Parameter Name="RiskDate01" />
            <asp:Parameter Name="DoctorLevel01" />
            <asp:Parameter Name="CompanyLevel01" />
            <asp:Parameter Name="WorkHours01" />
            <asp:Parameter Name="WorkDays01" />
            <asp:Parameter Name="RiskDate02" />
            <asp:Parameter Name="DoctorLevel02" />
            <asp:Parameter Name="CompanyLevel02" />
            <asp:Parameter Name="WorkHours02" />
            <asp:Parameter Name="WorkDays02" />
            <asp:Parameter Name="RiskDate03" />
            <asp:Parameter Name="DoctorLevel03" />
            <asp:Parameter Name="CompanyLevel03" />
            <asp:Parameter Name="WorkHours03" />
            <asp:Parameter Name="WorkDays03" />
            <asp:Parameter Name="RiskDate04" />
            <asp:Parameter Name="DoctorLevel04" />
            <asp:Parameter Name="CompanyLevel04" />
            <asp:Parameter Name="WorkHours04" />
            <asp:Parameter Name="WorkDays04" />
            <asp:Parameter Name="RiskDate05" />
            <asp:Parameter Name="DoctorLevel05" />
            <asp:Parameter Name="CompanyLevel05" />
            <asp:Parameter Name="WorkHours05" />
            <asp:Parameter Name="WorkDays05" />
            <asp:Parameter Name="RiskDate06" />
            <asp:Parameter Name="DoctorLevel06" />
            <asp:Parameter Name="CompanyLevel06" />
            <asp:Parameter Name="WorkHours06" />
            <asp:Parameter Name="WorkDays06" />
            <asp:Parameter Name="RiskDate07" />
            <asp:Parameter Name="DoctorLevel07" />
            <asp:Parameter Name="CompanyLevel07" />
            <asp:Parameter Name="WorkHours07" />
            <asp:Parameter Name="WorkDays07" />
            <asp:Parameter Name="RiskDate08" />
            <asp:Parameter Name="DoctorLevel08" />
            <asp:Parameter Name="CompanyLevel08" />
            <asp:Parameter Name="WorkHours08" />
            <asp:Parameter Name="WorkDays08" />
            <asp:Parameter Name="RiskDate09" />
            <asp:Parameter Name="DoctorLevel09" />
            <asp:Parameter Name="CompanyLevel09" />
            <asp:Parameter Name="WorkHours09" />
            <asp:Parameter Name="WorkDays09" />
            <asp:Parameter Name="RiskDate10" />
            <asp:Parameter Name="DoctorLevel10" />
            <asp:Parameter Name="CompanyLevel10" />
            <asp:Parameter Name="WorkHours10" />
            <asp:Parameter Name="WorkDays10" />
            <asp:Parameter Name="RiskDate11" />
            <asp:Parameter Name="DoctorLevel11" />
            <asp:Parameter Name="CompanyLevel11" />
            <asp:Parameter Name="WorkHours11" />
            <asp:Parameter Name="WorkDays11" />
            <asp:Parameter Name="RiskDate12" />
            <asp:Parameter Name="DoctorLevel12" />
            <asp:Parameter Name="CompanyLevel12" />
            <asp:Parameter Name="WorkHours12" />
            <asp:Parameter Name="WorkDays12" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="BuDate" />
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="RiskDate01" />
            <asp:Parameter Name="DoctorLevel01" />
            <asp:Parameter Name="CompanyLevel01" />
            <asp:Parameter Name="WorkHours01" />
            <asp:Parameter Name="WorkDays01" />
            <asp:Parameter Name="RiskDate02" />
            <asp:Parameter Name="DoctorLevel02" />
            <asp:Parameter Name="CompanyLevel02" />
            <asp:Parameter Name="WorkHours02" />
            <asp:Parameter Name="WorkDays02" />
            <asp:Parameter Name="RiskDate03" />
            <asp:Parameter Name="DoctorLevel03" />
            <asp:Parameter Name="CompanyLevel03" />
            <asp:Parameter Name="WorkHours03" />
            <asp:Parameter Name="WorkDays03" />
            <asp:Parameter Name="RiskDate04" />
            <asp:Parameter Name="DoctorLevel04" />
            <asp:Parameter Name="CompanyLevel04" />
            <asp:Parameter Name="WorkHours04" />
            <asp:Parameter Name="WorkDays04" />
            <asp:Parameter Name="RiskDate05" />
            <asp:Parameter Name="DoctorLevel05" />
            <asp:Parameter Name="CompanyLevel05" />
            <asp:Parameter Name="WorkHours05" />
            <asp:Parameter Name="WorkDays05" />
            <asp:Parameter Name="RiskDate06" />
            <asp:Parameter Name="DoctorLevel06" />
            <asp:Parameter Name="CompanyLevel06" />
            <asp:Parameter Name="WorkHours06" />
            <asp:Parameter Name="WorkDays06" />
            <asp:Parameter Name="RiskDate07" />
            <asp:Parameter Name="DoctorLevel07" />
            <asp:Parameter Name="CompanyLevel07" />
            <asp:Parameter Name="WorkHours07" />
            <asp:Parameter Name="WorkDays07" />
            <asp:Parameter Name="RiskDate08" />
            <asp:Parameter Name="DoctorLevel08" />
            <asp:Parameter Name="CompanyLevel08" />
            <asp:Parameter Name="WorkHours08" />
            <asp:Parameter Name="WorkDays08" />
            <asp:Parameter Name="RiskDate09" />
            <asp:Parameter Name="DoctorLevel09" />
            <asp:Parameter Name="CompanyLevel09" />
            <asp:Parameter Name="WorkHours09" />
            <asp:Parameter Name="WorkDays09" />
            <asp:Parameter Name="RiskDate10" />
            <asp:Parameter Name="DoctorLevel10" />
            <asp:Parameter Name="CompanyLevel10" />
            <asp:Parameter Name="WorkHours10" />
            <asp:Parameter Name="WorkDays10" />
            <asp:Parameter Name="RiskDate11" />
            <asp:Parameter Name="DoctorLevel11" />
            <asp:Parameter Name="CompanyLevel11" />
            <asp:Parameter Name="WorkHours11" />
            <asp:Parameter Name="WorkDays11" />
            <asp:Parameter Name="RiskDate12" />
            <asp:Parameter Name="DoctorLevel12" />
            <asp:Parameter Name="CompanyLevel12" />
            <asp:Parameter Name="WorkHours12" />
            <asp:Parameter Name="WorkDays12" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="IndexNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
</asp:Content>
