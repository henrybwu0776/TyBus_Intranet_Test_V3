<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="AnecdoteCase.aspx.cs" Inherits="TyBus_Intranet_Test_V3.AnecdoteCase" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="AnecdoteCaseForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">肇事案件處理</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:CheckBox ID="chkHasInsurance" runat="server" CssClass="text-Left-Black" Text="已出險" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:CheckBox ID="chkCaseClose" runat="server" CssClass="text-Left-Black" Text="已和解" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別（編號）：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eDepNo_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="eDepNo_Split_Search" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eDepNo_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbCaseDate_Search" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eCaseDate_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="eCaseDate_Split_Search" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eCaseDate_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbCarID_Search" runat="server" CssClass="text-Right-Blue" Text="牌照號碼" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eCarID_Search" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbDriver_Search" runat="server" CssClass="text-Right-Blue" Text="駕駛員" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eDriver_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eDriverName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColWidth-7Col" colspan="2">
                    <asp:FileUpload ID="fuExcel" runat="server" />
                    <asp:Button ID="bbImportData" Text="匯入資料" CssClass="button-Black" runat="server" OnClick="bbImportData_Click" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbSearch" Text="預覽" CssClass="button-Black" runat="server" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbExcel" Text="匯出EXCEL" CssClass="button-Blue" runat="server" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbCancel" Text="結束" CssClass="button-Red" runat="server" OnClick="bbCancel_Click" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plMainDataShow" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridAnecdoteCaseA_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4"
            DataKeyNames="CaseNo" DataSourceID="sdsAnecdoteCaseA_List" ForeColor="#333333" GridLines="None" PageSize="5" Width="100%">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="CaseNo" HeaderText="序號" ReadOnly="True" SortExpression="CaseNo" />
                <asp:CheckBoxField DataField="HasInsurance" HeaderText="出險" SortExpression="HasInsurance" />
                <asp:CheckBoxField DataField="CaseClose" HeaderText="和解" SortExpression="CaseClose" />
                <asp:BoundField DataField="DepName" HeaderText="站別" SortExpression="DepName" />
                <asp:BoundField DataField="BuildDate" DataFormatString="{0:d}" HeaderText="發生日期" SortExpression="BuildDate" />
                <asp:BoundField DataField="BuildManName" HeaderText="建檔人" ReadOnly="True" SortExpression="BuildManName" />
                <asp:BoundField DataField="Car_ID" HeaderText="牌照號碼" SortExpression="Car_ID" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員" SortExpression="DriverName" />
                <asp:BoundField DataField="InsuMan" HeaderText="保險經辦人" SortExpression="InsuMan" />
                <asp:BoundField DataField="AnecdotalResRatio" HeaderText="肇責比率" SortExpression="AnecdotalResRatio" />
                <asp:CheckBoxField DataField="IsNoDeduction" HeaderText="免扣精勤" SortExpression="IsNoDeduction" />
                <asp:BoundField DataField="DeductionDate" DataFormatString="{0:d}" HeaderText="扣發日期" SortExpression="DeductionDate" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" />
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
        <asp:FormView ID="fvAnecdoteCaseA_Data" runat="server" DataKeyNames="CaseNo" DataSourceID="sdsAnecdoteCaseA_Data" Width="100%" OnDataBound="fvAnecdoteCaseA_Data_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_Edit_Click" Text="確定" Width="90px" UseSubmitBehavior="False" OnClientClick="this.disabled=true;" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="90px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseNo_Edit" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eHasInsurance_Edit" runat="server" CssClass="text-Left-Black" Text="出險" AutoPostBack="true" OnCheckedChanged="eHasInsurance_Edit_CheckedChanged" />
                                    <asp:Label ID="eHasInsuranceTitle_Edit" runat="server" Text='<%# Eval("HasInsurance") %>' Visible="false" />
                                    <asp:CheckBox ID="eCaseClose_Edit" runat="server" CssClass="text-Left-Black" Text="和解" AutoPostBack="true" OnCheckedChanged="eCaseClose_Edit_CheckedChanged" />
                                    <asp:Label ID="eCaseCloseTitle_Edit" runat="server" Text='<%#Eval("CaseClose") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildDate_Edit" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBuildDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBulidMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="30%" />
                                    <asp:Label ID="eBuildManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildManName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Text='<%# Eval("DepNo") %>' Width="30%" />
                                    <asp:Label ID="edepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarID_Edit" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCarID_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDriver_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Edit_TextChanged" Text='<%# Eval("Driver") %>' Width="30%" />
                                    <asp:Label ID="eDriverName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbInsuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="保險經辦：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eInsuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InsuMan") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAnecdotalResRatio_Edit" runat="server" CssClass="text-Right-Blue" Text="肇責比率：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAnecdotalResRatio_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AnecdotalResRatio") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eIsNoDeduction_Edit" runat="server" CssClass="text-Left-Black" Text="免扣精勤" AutoPostBack="true" OnCheckedChanged="eIsNoDeduction_Edit_CheckedChanged" />
                                    <asp:Label ID="eIsNoDeductionTitle_Edit" runat="server" Text='<%# Eval("IsNoDeduction") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDeductionDate_Edit" runat="server" CssClass="text-Right-Blue" Text="精勤扣款日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDeductionDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DeductionDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" />
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseOccurrence_Edit" runat="server" CssClass="text-Right-Blue" Text="肇事經過：" Width="90%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eCaseOccurrence_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("CaseOccurrence") %>' Width="95%" Height="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="95%" Height="97%" />
                                </td>
                            </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                            <asp:Label ID="lbSubTitle_Edit" runat="server" CssClass="titleText-S-Blue" Text="鑑定會審議結果" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="cbIsExemption_Edit" runat="server" CssClass="text-Left-Black" Text="裁定免責" AutoPostBack="true" OnCheckedChanged="cbIsExemption_Edit_CheckedChanged" />
                            <asp:Label ID="eIsExemption_Edit" runat="server" Text='<%# Eval("IsExemption") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPaidAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="已自付總額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="ePaidAmount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PaidAmount") %>' Width="95%" AutoPostBack="true" OnTextChanged="ePaidAmount_Edit_TextChanged" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbInsuAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="保險理賠金" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eInsuAmount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InsuAmount") %>' Width="95%" AutoPostBack="true" OnTextChanged="ePaidAmount_Edit_TextChanged" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPenalty_Edit" runat="server" CssClass="text-Right-Blue" Text="罰款比例及分擔額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="ePenaltyRatio_Edit" runat="server" CssClass="text-Right-Black" Text='<%# Eval("PenaltyRatio") %>' Width="30%" AutoPostBack="true" OnTextChanged="ePaidAmount_Edit_TextChanged" />
                            <asp:Label ID="lbSplit_Edit" runat="server" CssClass="text-Left-Black" Text="％" Width="15%" />
                            <asp:Label ID="ePenalty_Edit" runat="server" CssClass="text-Left-Black" Text='<%#Eval("Penalty") %>' Width="45%" /> 
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
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CssClass="button-Black" CausesValidation="false" CommandName="New" Text="新增肇事單" />
            </EmptyDataTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_INS_Click" Text="確定" Width="90px" UseSubmitBehavior="False" OnClientClick="this.disabled=true;" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="90px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseNo_INS" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_INS" runat="server" CssClass="text-Left-Black" Text='<%#Eval("CaseNo") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eHasInsurance_INS" runat="server" CssClass="text-Left-Black" Text="出險" AutoPostBack="true" OnCheckedChanged="eHasInsurance_INS_CheckedChanged" />
                                    <asp:Label ID="eHasInsuranceTitle_INS" runat="server" Visible="false" />
                                    <asp:CheckBox ID="eCaseClose_INS" runat="server" CssClass="text-Left-Black" Text="和解" AutoPostBack="true" OnCheckedChanged="eCaseClose_INS_CheckedChanged" />
                                    <asp:Label ID="eCaseCloseTitle_INS" runat="server" Text='<%#Eval("CaseClose") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildDate_INS" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eBuildDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuildMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="30%" />
                                    <asp:Label ID="eBuildManNaMe_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildManName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Eval("DepNo") %>' Width="30%" />
                                    <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarID_INS" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCarID_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_INS" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDriver_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_INS_TextChanged" Text='<%# Eval("Driver") %>' Width="30%" />
                                    <asp:Label ID="eDriverName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbInsuMan_INS" runat="server" CssClass="text-Right-Blue" Text="保險經辦：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eInsuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InsuMan") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAnecdotalResRatio_INS" runat="server" CssClass="text-Right-Blue" Text="肇責比率：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAnecdotalResRatio_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AnecdotalResRatio") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eIsNoDeduction_INS" runat="server" CssClass="text-Left-Black" Text="免扣精勤" AutoPostBack="true" OnCheckedChanged="eIsNoDeduction_INS_CheckedChanged" />
                                    <asp:Label ID="eIsNoDeductionTitle_INS" runat="server" Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDeductionDate_INS" runat="server" CssClass="text-Right-Blue" Text="精勤扣款日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDeductionDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DeductionDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" />
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseOccurrence_INS" runat="server" CssClass="text-Right-Blue" Text="肇事經過：" Width="90%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eCaseOccurrence_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("CaseOccurrence") %>' Width="95%" Height="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="95%" Height="97%" />
                                </td>
                            </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                            <asp:Label ID="lbSubTitle_INS" runat="server" CssClass="titleText-S-Blue" Text="鑑定會審議結果" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="cbIsExemption_INS" runat="server" CssClass="text-Left-Black" Text="裁定免責" AutoPostBack="true" OnCheckedChanged="cbIsExemption_INS_CheckedChanged" />
                            <asp:Label ID="eIsExemption_INS" runat="server" Text='<%# Eval("IsExemption") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPaidAmount_INS" runat="server" CssClass="text-Right-Blue" Text="已自付總額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="ePaidAmount_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PaidAmount") %>' Width="95%" AutoPostBack="true" OnTextChanged="ePaidAmount_INS_TextChanged" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbInsuAmount_INS" runat="server" CssClass="text-Right-Blue" Text="保險理賠金" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:TextBox ID="eInsuAmount_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InsuAmount") %>' Width="95%" AutoPostBack="true" OnTextChanged="ePaidAmount_INS_TextChanged" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPenalty_INS" runat="server" CssClass="text-Right-Blue" Text="罰款比例及分擔額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="ePenaltyRatio_INS" runat="server" CssClass="text-Right-Black" Text='<%# Eval("PenaltyRatio") %>' Width="30%" AutoPostBack="true" OnTextChanged="ePaidAmount_INS_TextChanged" />
                            <asp:Label ID="lbSplit_INS" runat="server" CssClass="text-Left-Black" Text="％" Width="15%" />
                            <asp:Label ID="ePenalty_INS" runat="server" CssClass="text-Left-Black" Text='<%#Eval("Penalty") %>' Width="45%" /> 
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
            <ItemTemplate>
                <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CommandName="New" CssClass="button-Black" Text="新增肇事單" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CommandName="Edit" CssClass="button-Blue" Text="修改" Width="90px" />
                &nbsp;<asp:Button ID="bbDelete_Edit" runat="server" CausesValidation="False" OnClick="bbDelete_Edit_Click" CssClass="button-Red" Text="刪除" Width="90px" />
                &nbsp;<asp:Button ID="bbExportERP_List" runat="server" CausesValidation="false" OnClick="bbExportERP_Click" CssClass="button-Blue" Text="ERP同步" Width="120px" />
                &nbsp;<asp:Button ID="bbDelERP_List" runat="server" CausesValidation="false" OnClick="bbDelERP_List_Click" CssClass="button-Red" Text="取消同步" Width="120px" />
                &nbsp;<asp:Button ID="bbReport" runat="server" CausesValidation="false" OnClick="bbReport_Click" CssClass="button-Black" Text="列印審議通知" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseNo_List" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCaseNo_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("CaseNo") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eHasInsurance_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="出險" Checked='<%# Eval("HasInsurance") %>' />
                            <asp:CheckBox ID="eCaseClose_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="和解" Checked='<%# Eval("CaseClose") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildDate_List" runat="server" CssClass="text-Right-Blue" Text="發生日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuildDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuildMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="30%" />
                            <asp:Label ID="eBuildManNaMe_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildManName") %>' Width="55%" />
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
                            <asp:Label ID="lbCarID_List" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCarID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriver_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDriver_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver") %>' Width="30%" />
                            <asp:Label ID="eDriverName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbInsuMan_List" runat="server" CssClass="text-Right-Blue" Text="保險經辦：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eInsuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InsuMan") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAnecdotalResRatio_List" runat="server" CssClass="text-Right-Blue" Text="肇責比率：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAnecdotalResRatio_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AnecdotalResRatio") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eIsNoDeduction_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="免扣精勤" Checked='<%# Eval("IsNoDeduction") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDeductionDate_List" runat="server" CssClass="text-Right-Blue" Text="精勤扣款日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDeductionDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DeductionDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" >
                            <asp:Label ID="eERPCouseNo_List" runat="server" Text='<%# Eval("ERPCouseNo") %>' />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseOccurrence_List" runat="server" CssClass="text-Right-Blue" Text="肇事經過：" Width="90%" />
                        </td>
                        <td class="MultiLine_Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eCaseOccurrence_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("CaseOccurrence") %>' Width="95%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                        </td>
                        <td class="MultiLine_Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Width="95%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                            <asp:Label ID="lbSubTitle_List" runat="server" CssClass="titleText-S-Blue" Text="鑑定會審議結果" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="cbIsExemption_List" runat="server" CssClass="text-Left-Black" Text="裁定免責" />
                            <asp:Label ID="eIsExemption_List" runat="server" Text='<%# Eval("IsExemption") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPaidAmount_List" runat="server" CssClass="text-Right-Blue" Text="已自付總額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="ePaidAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PaidAmount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbInsuAmount_List" runat="server" CssClass="text-Right-Blue" Text="保險理賠金" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eInsuAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InsuAmount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbPenalty_List" runat="server" CssClass="text-Right-Blue" Text="罰款比例及分擔額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="ePenaltyRatio_List" runat="server" CssClass="text-Right-Black" Text='<%# Eval("PenaltyRatio") %>' Width="30%" />
                            <asp:Label ID="lbSplit_List" runat="server" CssClass="text-Left-Black" Text="％" Width="15%" />
                            <asp:Label ID="ePenalty_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("Penalty") %>' Width="45%" /> 
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
    <asp:SqlDataSource ID="sdsAnecdoteCaseA_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        SelectCommand="SELECT CaseNo, HasInsurance, DepName, BuildDate, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuildMan)) AS BuildManName, Car_ID, DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, ERPCouseNo, CaseClose, IsExemption, PaidAmount, InsuAmount, Penalty, PenaltyRatio FROM AnecdoteCase AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsAnecdoteCaseA_Data" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" 
        DeleteCommand="DELETE FROM AnecdoteCase WHERE (CaseNo = @CaseNo)" 
        InsertCommand="INSERT INTO AnecdoteCase(CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, CaseOccurrence,CaseClose) VALUES (@CaseNo, @HasInsurance, @DepNo, @DepName, @BuildDate, @BuildMan, @Car_ID, @Driver, @DriverName, @InsuMan, @AnecdotalResRatio, @IsNoDeduction, @DeductionDate, @Remark, @CaseOccurrence,@CaseClose)" 
        SelectCommand="SELECT CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuildMan)) AS BuildManName, Car_ID, Driver, DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose, InsuAmount, PenaltyRatio, Penalty, PaidAmount, IsExemption FROM AnecdoteCase AS a WHERE (CaseNo = @CaseNo)" 
        UpdateCommand="UPDATE AnecdoteCase SET HasInsurance = @HasInsurance, DepNo = @DepNo, DepName = @DepName, Car_ID = @Car_ID, Driver = @Driver, DriverName = @DriverName, InsuMan = @InsuMan, AnecdotalResRatio = @AnecdotalResRatio, IsNoDeduction = @IsNoDeduction, DeductionDate = @DeductionDate, Remark = @Remark, BuildDate = @BuildDate, CaseOccurrence = @CaseOccurrence,CaseClose=@CaseClose WHERE (CaseNo = @CaseNo)">
        <DeleteParameters>
            <asp:Parameter Name="CaseNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="HasInsurance" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="DepName" />
            <asp:Parameter Name="BuildDate" />
            <asp:Parameter Name="BuildMan" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="DriverName" />
            <asp:Parameter Name="InsuMan" />
            <asp:Parameter Name="AnecdotalResRatio" />
            <asp:Parameter Name="IsNoDeduction" />
            <asp:Parameter Name="DeductionDate" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="CaseOccurrence" />
            <asp:Parameter Name="CaseClose" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridAnecdoteCaseA_List" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="HasInsurance" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="DepName" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="DriverName" />
            <asp:Parameter Name="InsuMan" />
            <asp:Parameter Name="AnecdotalResRatio" />
            <asp:Parameter Name="IsNoDeduction" />
            <asp:Parameter Name="DeductionDate" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuildDate" />
            <asp:Parameter Name="CaseOccurrence" />
            <asp:Parameter Name="CaseClose" />
            <asp:Parameter Name="CaseNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:Panel ID="plDetailDataShow" runat="server" CssClass="ShowPanel-Detail">
        <asp:GridView ID="gridAnecdoteCaseB_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="CaseNoItems" DataSourceID="sdsAnecdoteCaseB_List" ForeColor="#333333" GridLines="None" PageSize="5" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="CaseNo" HeaderText="CaseNo" SortExpression="CaseNo" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                <asp:BoundField DataField="CaseNoItems" HeaderText="CaseNoItems" ReadOnly="True" SortExpression="CaseNoItems" Visible="False" />
                <asp:BoundField DataField="Relationship" HeaderText="對方姓名" SortExpression="Relationship" />
                <asp:BoundField DataField="RelCar_ID" HeaderText="對方車號" SortExpression="RelCar_ID" />
                <asp:BoundField DataField="EstimatedAmount" HeaderText="預估金額" SortExpression="EstimatedAmount" />
                <asp:BoundField DataField="ThirdInsurance" HeaderText="第三責任險" SortExpression="ThirdInsurance" />
                <asp:BoundField DataField="CompInsurance" HeaderText="強制險" SortExpression="CompInsurance" />
                <asp:BoundField DataField="PassengerInsu" HeaderText="乘客險" SortExpression="PassengerInsu" />
                <asp:BoundField DataField="DriverSharing" HeaderText="駕駛員負擔金額" SortExpression="DriverSharing" />
                <asp:BoundField DataField="CompanySharing" HeaderText="公司負擔金額" SortExpression="CompanySharing" />
                <asp:BoundField DataField="CarDamageAMT" HeaderText="車損金額" SortExpression="CarDamageAMT" />
                <asp:BoundField DataField="PersonDamageAMT" HeaderText="體傷金額" SortExpression="PersonDamageAMT" />
                <asp:BoundField DataField="RelationComp" HeaderText="對方賠付金額" SortExpression="RelationComp" />
                <asp:BoundField DataField="ReconciliationDate" DataFormatString="{0:d}" HeaderText="和解日期" SortExpression="ReconciliationDate" />
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
        <asp:FormView ID="fvAnecdoteCaseB_Data" runat="server" DataKeyNames="CaseNoItems" DataSourceID="sdsAnecdoteCaseB_Data" Width="100%" OnDataBound="fvAnecdoteCaseB_Data_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOKCaseB_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKCaseB_Edit_Click" Text="確定" Width="90px" />
                &nbsp;<asp:Button ID="bbCancelCaseB_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="90px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBItems_Edit" runat="server" CssClass="text-Right-Blue" Text="單號項次：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                            <asp:Label ID="eCaseBCaseNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="55%" />
                            <asp:Label ID="eCaseBItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="30%" />
                            <asp:Label ID="eCaseBCaseNoItems_Edit" runat="server" Text='<%# Eval("CaseNoItems") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBRelationship_Edit" runat="server" CssClass="text-Right-Blue" Text="對方姓名：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBRelationship_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Relationship") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBRelCar_ID_Edit" runat="server" CssClass="text-Right-Blue" Text="對方車號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBRelCarID_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelCar_ID") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBEstimatedAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="預估金額：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBEstimatedAmount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EstimatedAmount") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBThirdInsurance_Edit" runat="server" CssClass="text-Right-Blue" Text="第三責任險：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBThirdInsurance_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ThirdInsurance") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBCompInsurance_Edit" runat="server" CssClass="text-Right-Blue" Text="強制險：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBCompInsurance_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompInsurance") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBPassengerInsu_Edit" runat="server" CssClass="text-Right-Blue" Text="乘客險：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBPassengerInsu_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PassengerInsu") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBDriverSharing_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛員負擔：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBDriverSharing_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverSharing") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBCompanySharing_Edit" runat="server" CssClass="text-Right-Blue" Text="公司負擔：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBCompanySharing_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompanySharing") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBRelationComp_Edit" runat="server" CssClass="text-Right-Blue" Text="對方賠付：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBRelationComp_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelationComp") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBCarDamageAMT_Edit" runat="server" CssClass="text-Right-Blue" Text="車損金額：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBCarDamageAMT_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarDamageAMT") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBPersonDamageAMT_Edit" runat="server" CssClass="text-Right-Blue" Text="體傷金額：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBPersonDamageAMT_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PersonDamageAMT") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBReconciliationDate_Edit" runat="server" CssClass="text-Right-Blue" Text="和解日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBReconciliationDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ReconciliationDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                        </td>
                        <td class="MultiLine_Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eCaseBRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="97%" Height="95%" />
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
            </EditItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNewCaseB_Empty" runat="server" CausesValidation="false" CssClass="button-Black" CommandName="New" Text="新增肇事明細" />
            </EmptyDataTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOKCaseB_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOKCaseB_INS_Click" Text="確定" Width="90px" />
                &nbsp;<asp:Button ID="bbCancelCaseB_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="90px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBItems_INS" runat="server" CssClass="text-Right-Blue" Text="單號項次：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                            <asp:Label ID="eCaseBCaseNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="55%" />
                            <asp:Label ID="eCaseBItems_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="30%" />
                            <asp:Label ID="eCaseBCaseNoItems_INS" runat="server" Text='<%# Eval("CaseNoItems") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBRelationship_INS" runat="server" CssClass="text-Right-Blue" Text="對方姓名：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBRelationship_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Relationship") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBRelCar_ID_INS" runat="server" CssClass="text-Right-Blue" Text="對方車號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBRelCarID_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelCar_ID") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBEstimatedAmount_INS" runat="server" CssClass="text-Right-Blue" Text="預估金額：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBEstimatedAmount_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EstimatedAmount") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBThirdInsurance_INS" runat="server" CssClass="text-Right-Blue" Text="第三責任險：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBThirdInsurance_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ThirdInsurance") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBCompInsurance_INS" runat="server" CssClass="text-Right-Blue" Text="強制險：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBCompInsurance_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompInsurance") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBPassengerInsu_INS" runat="server" CssClass="text-Right-Blue" Text="乘客險：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBPassengerInsu_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PassengerInsu") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBDriverSharing_INS" runat="server" CssClass="text-Right-Blue" Text="駕駛員負擔：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBDriverSharing_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverSharing") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBCompanySharing_INS" runat="server" CssClass="text-Right-Blue" Text="公司負擔：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBCompanySharing_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompanySharing") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBRelationComp_INS" runat="server" CssClass="text-Right-Blue" Text="對方賠付：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBRelationComp_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelationComp") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBCarDamageAMT_INS" runat="server" CssClass="text-Right-Blue" Text="車損金額：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBCarDamageAMT_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarDamageAMT") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBPersonDamageAMT_INS" runat="server" CssClass="text-Right-Blue" Text="體傷金額：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBPersonDamageAMT_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PersonDamageAMT") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBReconciliationDate_INS" runat="server" CssClass="text-Right-Blue" Text="和解日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:TextBox ID="eCaseBReconciliationDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ReconciliationDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                        </td>
                        <td class="MultiLine_Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eCaseBRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="97%" Height="95%" />
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
            </InsertItemTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNewCaseB_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增肇事明細" />
                &nbsp;<asp:Button ID="bbEditCaseB_List" runat="server" CausesValidation="False" CssClass="button-Blue" CommandName="Edit" Text="修改" Width="90px" />
                &nbsp;<asp:Button ID="bbDelCaseB_List" runat="server" CausesValidation="False" CssClass="button-Red" OnClick="bbDelCaseB_List_Click" Text="刪除" Width="90px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBItems_List" runat="server" CssClass="text-Right-Blue" Text="單號項次：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                            <asp:Label ID="eCaseBCaseNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseNo") %>' Width="55%" />
                            <asp:Label ID="eCaseBItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="30%" />
                            <asp:Label ID="eCaseBCaseNoItems_List" runat="server" Text='<%# Eval("CaseNoItems") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBRelationship_List" runat="server" CssClass="text-Right-Blue" Text="對方姓名：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseBRelationship_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Relationship") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBRelCar_ID_List" runat="server" CssClass="text-Right-Blue" Text="對方車號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseBRelCarID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelCar_ID") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBEstimatedAmount_List" runat="server" CssClass="text-Right-Blue" Text="預估金額：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseBEstimatedAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EstimatedAmount") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBThirdInsurance_List" runat="server" CssClass="text-Right-Blue" Text="第三責任險：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseBThirdInsurance_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ThirdInsurance") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBCompInsurance_List" runat="server" CssClass="text-Right-Blue" Text="強制險：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseBCompInsurance_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompInsurance") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBPassengerInsu_List" runat="server" CssClass="text-Right-Blue" Text="乘客險：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseBPassengerInsu_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PassengerInsu") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBDriverSharing_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員負擔：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseBDriverSharing_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverSharing") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBCompanySharing_List" runat="server" CssClass="text-Right-Blue" Text="公司負擔：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseBCompanySharing_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompanySharing") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBRelationComp_List" runat="server" CssClass="text-Right-Blue" Text="對方賠付：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseBRelationComp_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelationComp") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBCarDamageAMT_List" runat="server" CssClass="text-Right-Blue" Text="車損金額：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseBCarDamageAMT_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarDamageAMT") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBPersonDamageAMT_List" runat="server" CssClass="text-Right-Blue" Text="體傷金額：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseBPersonDamageAMT_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PersonDamageAMT") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBReconciliationDate_List" runat="server" CssClass="text-Right-Blue" Text="和解日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCaseBReconciliationDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ReconciliationDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCaseBRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                        </td>
                        <td class="MultiLine_Low ColBorder ColWidth-6Col" colspan="5">
                            <asp:TextBox ID="eCaseBRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Width="97%" Height="95%" />
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
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
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
            <LocalReport ReportPath="Report\AnecdoteCaseP2.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsAnecdoteCaseB_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, PassengerInsu, Remark FROM AnecdoteCaseB WHERE (CaseNo = @CaseNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridAnecdoteCaseA_List" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsAnecdoteCaseB_Data" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" DeleteCommand="DELETE FROM AnecdoteCaseB WHERE (CaseNoItems = @CaseNoItems)" InsertCommand="INSERT INTO AnecdoteCaseB(CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, PassengerInsu, Remark) VALUES (@CaseNo, @Items, @CaseNoItems, @Relationship, @RelCar_ID, @EstimatedAmount, @ThirdInsurance, @CompInsurance, @DriverSharing, @CompanySharing, @CarDamageAMT, @PersonDamageAMT, @RelationComp, @ReconciliationDate, @PassengerInsu, @Remark)" SelectCommand="SELECT CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, PassengerInsu, Remark FROM AnecdoteCaseB WHERE (CaseNoItems = @CaseNoCaseNoItems)" UpdateCommand="UPDATE AnecdoteCaseB SET Relationship = @Relationship, RelCar_ID = @RelCar_ID, EstimatedAmount = @EstimatedAmount, ThirdInsurance = @ThirdInsurance, CompInsurance = @CompInsurance, DriverSharing = @DriverSharing, CompanySharing = @CompanySharing, CarDamageAMT = @CarDamageAMT, PersonDamageAMT = @PersonDamageAMT, RelationComp = @RelationComp, ReconciliationDate = @ReconciliationDate, PassengerInsu = @PassengerInsu, Remark = @Remark WHERE (CaseNoItems = @CaseNoItems)">
        <DeleteParameters>
            <asp:Parameter Name="CaseNoItems" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="Items" />
            <asp:Parameter Name="CaseNoItems" />
            <asp:Parameter Name="Relationship" />
            <asp:Parameter Name="RelCar_ID" />
            <asp:Parameter Name="EstimatedAmount" />
            <asp:Parameter Name="ThirdInsurance" />
            <asp:Parameter Name="CompInsurance" />
            <asp:Parameter Name="DriverSharing" />
            <asp:Parameter Name="CompanySharing" />
            <asp:Parameter Name="CarDamageAMT" />
            <asp:Parameter Name="PersonDamageAMT" />
            <asp:Parameter Name="RelationComp" />
            <asp:Parameter Name="ReconciliationDate" />
            <asp:Parameter Name="PassengerInsu" />
            <asp:Parameter Name="Remark" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridAnecdoteCaseB_List" Name="CaseNoCaseNoItems" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="Relationship" />
            <asp:Parameter Name="RelCar_ID" />
            <asp:Parameter Name="EstimatedAmount" />
            <asp:Parameter Name="ThirdInsurance" />
            <asp:Parameter Name="CompInsurance" />
            <asp:Parameter Name="DriverSharing" />
            <asp:Parameter Name="CompanySharing" />
            <asp:Parameter Name="CarDamageAMT" />
            <asp:Parameter Name="PersonDamageAMT" />
            <asp:Parameter Name="RelationComp" />
            <asp:Parameter Name="ReconciliationDate" />
            <asp:Parameter Name="PassengerInsu" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="CaseNoItems" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:Panel ID="plSHowDataC" runat="server" CssClass="ShowPanel-Detail">
        <asp:GridView ID="gridAnecdoteCaseC_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="CaseNoItem" DataSourceID="sdsAnecdoteCaseC_List" ForeColor="#333333" GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <Columns>
                <asp:BoundField DataField="CaseNoItem" HeaderText="CaseNoItem" ReadOnly="True" SortExpression="CaseNoItem" Visible="False" />
                <asp:BoundField DataField="CaseNo" HeaderText="CaseNo" SortExpression="CaseNo" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" SortExpression="Items" />
                <asp:BoundField DataField="ContactDate" DataFormatString="{0:d}" HeaderText="連絡日期" SortExpression="ContactDate" />
                <asp:BoundField DataField="Excutort_C" HeaderText="處理人員" ReadOnly="True" SortExpression="Excutort_C" />
                <asp:BoundField DataField="ContactPerson" HeaderText="連絡對象" SortExpression="ContactPerson" />
                <asp:BoundField DataField="ContactNote" HeaderText="處理概述" SortExpression="ContactNote" />
                <asp:BoundField DataField="AssignDate" DataFormatString="{0:d}" HeaderText="轉交日期" SortExpression="AssignDate" />
                <asp:BoundField DataField="AssignedMan_C" HeaderText="轉交人員" ReadOnly="True" SortExpression="AssignedMan_C" />
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
    </asp:Panel>
    <asp:SqlDataSource ID="sdsAnecdoteCaseC_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT CaseNoItem, CaseNo, Items, ContactPerson, ContactNote, ContactDate, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = B.Excutort)) AS Excutort_C, AssignDate, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_3 WHERE (EMPNO = B.AssignedMan)) AS AssignedMan_C FROM AnecdoteCaseC AS B WHERE (CaseNo = @CaseNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridAnecdoteCaseA_List" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
</asp:Content>
