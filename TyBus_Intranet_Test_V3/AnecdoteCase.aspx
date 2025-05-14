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
                    <asp:Label ID="lbCaseDate_Search" runat="server" CssClass="text-Right-Blue" Text="出險日期：" Width="90%" />
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
                <td class="ColWidth-7Col" colspan="4">
                    <asp:Label ID="eErrorMSG_Main" runat="server" CssClass="errorMessageText" Visible="false" />
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
        <div>
            <asp:Label ID="eErrorMSG_A" runat="server" CssClass="errorMessageText" Visible="false" />
        </div>
        <asp:GridView ID="gridAnecdoteCaseA_List" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            CellPadding="4" ForeColor="#333333" DataKeyNames="CaseNo" GridLines="None" PageSize="5" Width="100%" OnPageIndexChanging="gridAnecdoteCaseA_List_PageIndexChanging" DataSourceID="sdsAnecdoteCaseA_List">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="true" ShowCancelButton="False" />
                <asp:BoundField DataField="CaseNo" HeaderText="序號" ReadOnly="True" SortExpression="CaseNo" />
                <asp:CheckBoxField DataField="HasInsurance" HeaderText="出險" SortExpression="HasInsurance" />
                <asp:CheckBoxField DataField="CaseClose" HeaderText="和解" SortExpression="CaseClose" />
                <asp:BoundField DataField="DepName" HeaderText="站別" SortExpression="DepName" />
                <asp:BoundField DataField="BuildDate" DataFormatString="{0:d}" HeaderText="出險日期" SortExpression="BuildDate" />
                <asp:BoundField DataField="BuildManName" HeaderText="建檔人" ReadOnly="True" SortExpression="BuildManName" />
                <asp:BoundField DataField="Car_ID" HeaderText="牌照號碼" SortExpression="Car_ID" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員" SortExpression="DriverName" />
                <asp:BoundField DataField="InsuMan" HeaderText="保險經辦人" SortExpression="InsuMan" />
                <asp:BoundField DataField="AnecdotalResRatio" HeaderText="肇責比率" SortExpression="AnecdotalResRatio" />
                <asp:CheckBoxField DataField="IsNoDeduction" HeaderText="免扣精勤" SortExpression="IsNoDeduction" />
                <asp:BoundField DataField="DeductionDate" DataFormatString="{0:d}" HeaderText="扣發日期" SortExpression="DeductionDate" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" />
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
        <asp:FormView ID="fvAnecdoteCaseA_Data" runat="server" DataKeyNames="CaseNo" DataSourceID="sdsAnecdoteCaseA_Data" Width="100%" OnDataBound="fvAnecdoteCaseA_Data_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Black" OnClick="bbOK_Edit_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCaseNo_A_Edit" runat="server" CssClass="text-Right-Blue" Text="肇事單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eCaseNo_A_Edit" runat="server" Text='<%# Eval("CaseNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbHasInsurance_Edit" runat="server" CssClass="text-Right-Blue" Text="處理情況" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:CheckBox ID="cbHasInsurance_Edit" runat="server" CssClass="text-Left-Black" Text="出險" AutoPostBack="true" OnCheckedChanged="cbHasInsurance_Edit_CheckedChanged" Width="35%" />
                            <asp:CheckBox ID="cbCaseClose_Edit" runat="server" CssClass="text-Left-Black" Text="和解" AutoPostBack="true" OnCheckedChanged="cbCaseClose_Edit_CheckedChanged" Width="35%" />
                            <asp:Label ID="eHasInsurance_Edit" runat="server" Text='<%# Eval("HasInsurance") %>' Visible="false" />
                            <asp:Label ID="eCaseClose_Edit" runat="server" Text='<%# Eval("CaseClose") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuildDate_Edit" runat="server" CssClass="text-Right-Blue" Text="出險日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eBuildDate_Edit" runat="server" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuildMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eBuildMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' AutoPostBack="true" OnTextChanged="eBuildMan_Edit_TextChanged" Width="35%" />
                            <asp:Label ID="eBuildManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbReportDate_Edit" runat="server" CssClass="text-Right-Blue" Text="報告日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eReportDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ReportDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbInsuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="保險經辦" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eInsuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InsuMan") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAnecdotalResRatio_Edit" runat="server" CssClass="text-Right-Blue" Text="肇責比例" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eAnecdotalResRatio_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AnecdotalResRatio") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:CheckBox ID="cbIsNoDeduction_Edit" runat="server" CssClass="text-Left-Black" Text="免扣精勤" AutoPostBack="true" OnCheckedChanged="cbIsNoDeduction_Edit_CheckedChanged" Width="95%" />
                            <asp:Label ID="eIsNoDeduction_Edit" runat="server" Text='<%# Eval("IsNoDeduction") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDeductionDate_Edit" runat="server" CssClass="text-Right-Blue" Text="減發精勤日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eDeductionDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DeductionDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRemark_A_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                            <asp:TextBox ID="eRemark_A_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Remark_A") %>' TextMode="MultiLine" Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                            <asp:Label ID="lbSubTitle_Driver_Edit" runat="server" CssClass="titleText-Blue" Text="駕駛員資料" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCar_ID_Edit" runat="server" CssClass="text-Right-Blue" Text="牌照號碼" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eCar_ID_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eCar_ID_Edit_TextChanged" Text='<%# Eval("Car_ID") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDepName_Edit" runat="server" CssClass="text-Right-Blue" Text="所屬單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Width="35%" />
                            <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDriver_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eDriver_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver") %>' AutoPostBack="true" OnTextChanged="eDriver_Edit_TextChanged" Width="35%" />
                            <asp:Label ID="eDriverName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbIDCardNo_Edit" runat="server" CssClass="text-Right-Blue" Text="身分證字號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eIDCardNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("IDCardNo") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBirthday_Edit" runat="server" CssClass="text-Right-Blue" Text="出生日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eBirthday_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Birthday","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAssumeday_Edit" runat="server" CssClass="text-Right-Blue" Text="到職日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAssumeday_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Assumeday","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTelephoneNo_Edit" runat="server" CssClass="text-Right-Blue" Text="連絡電話" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eTelephoneNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TelephoneNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAddress_Edit" runat="server" CssClass="text-Right-Blue" Text="地址" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                            <asp:TextBox ID="eAddress_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Address") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPersonDamage_Edit" runat="server" CssClass="text-Right-Blue" Text="體傷" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="4">
                            <asp:TextBox ID="ePersonDamage_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PersonDamage") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCarDamage_Edit" runat="server" CssClass="text-Right-Blue" Text="車 (財) 損" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="4">
                            <asp:TextBox ID="eCarDamage_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarDamage") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                            <asp:Label ID="lbSubTitle_CaseNote_Edit" runat="server" CssClass="titleText-Blue" Text="事故相關訊息" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCaseDate_Edit" runat="server" CssClass="text-Right-Blue" Text="事故日期 / 時間" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eCaseDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="55%" />
                            <asp:TextBox ID="eCaseTime_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseTime") %>' Width="35%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbOutReportNo_Edit" runat="server" CssClass="text-Right-Blue" Text="登記聯單編號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eOutReportNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OutReportNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCasePosition_Edit" runat="server" CssClass="text-Right-Blue" Text="事故地點" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="4">
                            <asp:TextBox ID="eCasePosition_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CasePosition") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPoliceUnit_Edit" runat="server" CssClass="text-Right-Blue" Text="警方受理單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                            <asp:TextBox ID="ePoliceUnit_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PoliceUnit") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPoliceName_Edit" runat="server" CssClass="text-Right-Blue" Text="承辦員警" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePoliceName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PoliceName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:CheckBox ID="cbHasVideo_Edit" runat="server" CssClass="text-Left-Black" Text="有行車影像" AutoPostBack="true" OnCheckedChanged="cbHasVideo_Edit_CheckedChanged" Width="95%" />
                            <asp:Label ID="eHasVideo_Edit" runat="server" Text='<%# Eval("HasVideo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbNoVodeoReason_Edit" runat="server" CssClass="text-Right-Blue" Text="無影像原因" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eNoVideoReason_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("NoVideoReason") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCaseOccurrence_Edit" runat="server" CssClass="text-Right-Blue" Text="肇事經過" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                            <asp:TextBox ID="eCaseOccurrence_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseOccurrence") %>' Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbHasCaseData_Edit" runat="server" CssClass="text-Right-Blue" Text="已申請資料" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="6">
                            <asp:RadioButtonList ID="rbHasCaseData_Edit" runat="server" CssClass="text-Left-Black" RepeatColumns="1" AutoPostBack="true" OnSelectedIndexChanged="rbHasCaseData_Edit_SelectedIndexChanged">
                                <asp:ListItem Text="有" Value="Y" Selected="True" />
                                <asp:ListItem Text="無申請初判表、現場圖、事故現場照片" Value="N" />
                            </asp:RadioButtonList>
                            <asp:Label ID="eHasCaseData_Edit" runat="server" Text='<%# Eval("HasCaseData") %>' Visible="false" />
                            <br />
                            <asp:CheckBox ID="cbHasAccReport_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbHasAccReport_Edit_CheckedChanged" Text="已申請行車事故鑑定" />
                            <asp:Label ID="eHasAccReport_Edit" runat="server" Text='<%# Eval("HasAccReport") %>' Visible="false" />
                            <asp:Label ID="eERPCouseNo_Edit" runat="server" Text='<%# Eval("ERPCouseNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動資料" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyMan_A_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_A") %>' Width="35%" />
                            <asp:Label ID="eModifyManName_A_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName_A") %>' Width="60%" />
                            <br />
                            <asp:Label ID="eModifyDate_A_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate_A","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                            <asp:Label ID="lbSubTitle_Edit" runat="server" CssClass="titleText-Blue" Text="鑑定會審議結果" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbIsExemption_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛責任" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:CheckBox ID="cbIsExemption_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbIsExemption_Edit_CheckedChanged" Text="裁定免責" />
                            <asp:Label ID="eIsExemption_Edit" runat="server" Text='<%# Eval("IsExemption") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPaidAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="已自付總額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePaidAmount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PaidAmount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbInsuAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="保險理賠金" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eInsuAmount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InsuAmount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPenaltyRatio_Edit" runat="server" CssClass="text-Right-Blue" Text="罰款比例" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePenaltyRatio_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PenaltyRatio") %>' Width="75%" />
                            <a>％</a>
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPenalty_Edit" runat="server" CssClass="text-Right-Blue" Text="罰款分擔額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePenalty_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Penalty") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                    </tr>
                </table>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Black" OnClick="bbOK_INS_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCaseNo_A_INS" runat="server" CssClass="text-Right-Blue" Text="肇事單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eCaseNo_A_INS" runat="server" Text='<%# Eval("CaseNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbHasInsurance_INS" runat="server" CssClass="text-Right-Blue" Text="處理情況" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:CheckBox ID="cbHasInsurance_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbHasInsurance_INS_CheckedChanged" Text="出險" Width="35%" />
                            <asp:CheckBox ID="cbCaseClose_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbCaseClose_INS_CheckedChanged" Text="和解" Width="35%" />
                            <asp:Label ID="eHasInsurance_INS" runat="server" Text='<%# Eval("HasInsurance") %>' Visible="false" />
                            <asp:Label ID="eCaseClose_INS" runat="server" Text='<%# Eval("CaseClose") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuildDate_INS" runat="server" CssClass="text-Right-Blue" Text="出險日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eBuildDate_INS" runat="server" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuildMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eBuildMan_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eBuildMan_INS_TextChanged" Text='<%# Eval("BuildMan") %>' Width="35%" />
                            <asp:Label ID="eBuildManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbReportDate_INS" runat="server" CssClass="text-Right-Blue" Text="報告日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eReportDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ReportDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbInsuMan_INS" runat="server" CssClass="text-Right-Blue" Text="保險經辦" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eInsuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InsuMan") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAnecdotalResRatio_INS" runat="server" CssClass="text-Right-Blue" Text="肇責比例" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eAnecdotalResRatio_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AnecdotalResRatio") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:CheckBox ID="cbIsNoDeduction_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbIsNoDeduction_INS_CheckedChanged" Text="免扣精勤" />
                            <asp:Label ID="eIsNoDeduction_INS" runat="server" Text='<%# Eval("IsNoDeduction") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDeductionDate_INS" runat="server" CssClass="text-Right-Blue" Text="減發精勤日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eDeductionDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DeductionDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRemark_A_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                            <asp:TextBox ID="eRemark_A_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Remark_A") %>' TextMode="MultiLine" Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                            <asp:Label ID="lbSubTitle_Driver_INS" runat="server" CssClass="titleText-Blue" Text="駕駛員資料" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCar_ID_INS" runat="server" CssClass="text-Right-Blue" Text="牌照號碼" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eCar_ID_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eCar_ID_INS_TextChanged" Text='<%# Eval("Car_ID") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDepName_INS" runat="server" CssClass="text-Right-Blue" Text="所屬單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDriver_INS" runat="server" CssClass="text-Right-Blue" Text="駕駛員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eDriver_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_INS_TextChanged" Text='<%# Eval("Driver") %>' Width="35%" />
                            <asp:Label ID="eDriverName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbIDCardNo_INS" runat="server" CssClass="text-Right-Blue" Text="身分證字號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eIDCardNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("IDCardNo") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBirthday_INS" runat="server" CssClass="text-Right-Blue" Text="出生日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eBirthday_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Birthday","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAssumeday_INS" runat="server" CssClass="text-Right-Blue" Text="到職日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAssumeday_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Assumeday","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTelephoneNo_INS" runat="server" CssClass="text-Right-Blue" Text="連絡電話" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eTelephoneNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TelephoneNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAddress_INS" runat="server" CssClass="text-Right-Blue" Text="地址" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                            <asp:TextBox ID="eAddress_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Address") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPersonDamage_INS" runat="server" CssClass="text-Right-Blue" Text="體傷" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="4">
                            <asp:TextBox ID="ePersonDamage_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PersonDamage") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCarDamage_INS" runat="server" CssClass="text-Right-Blue" Text="車 (財) 損" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="4">
                            <asp:TextBox ID="eCarDamage_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarDamage") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                            <asp:Label ID="lbSubTitle_CaseNote_INS" runat="server" CssClass="titleText-Blue" Text="事故相關訊息" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCaseDate_INS" runat="server" CssClass="text-Right-Blue" Text="事故日期 / 時間" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eCaseDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="55%" />
                            <asp:TextBox ID="eCaseTime_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseTime") %>' Width="35%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbOutReportNo_INS" runat="server" CssClass="text-Right-Blue" Text="登記聯單編號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eOutReportNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OutReportNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCasePosition_INS" runat="server" CssClass="text-Right-Blue" Text="事故地點" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="4">
                            <asp:TextBox ID="eCasePosition_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CasePosition") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPoliceUnit_INS" runat="server" CssClass="text-Right-Blue" Text="警方受理單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                            <asp:TextBox ID="ePoliceUnit_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PoliceUnit") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPoliceName_INS" runat="server" CssClass="text-Right-Blue" Text="承辦員警" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePoliceName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PoliceName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:CheckBox ID="cbHasVideo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbHasVideo_INS_CheckedChanged" Text="有行車影像" Width="95%" />
                            <asp:Label ID="eHasVideo_INS" runat="server" Text='<%# Eval("HasVideo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbNoVodeoReason_INS" runat="server" CssClass="text-Right-Blue" Text="無影像原因" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:TextBox ID="eNoVideoReason_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("NoVideoReason") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCaseOccurrence_INS" runat="server" CssClass="text-Right-Blue" Text="肇事經過" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                            <asp:TextBox ID="eCaseOccurrence_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseOccurrence") %>' Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbHasCaseData_INS" runat="server" CssClass="text-Right-Blue" Text="已申請資料" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="6">
                            <asp:RadioButtonList ID="rbHasCaseData_INS" runat="server" CssClass="text-Left-Black" RepeatColumns="1" AutoPostBack="true" OnSelectedIndexChanged="rbHasCaseData_INS_SelectedIndexChanged">
                                <asp:ListItem Text="有" Value="V" Selected="True" />
                                <asp:ListItem Text="無申請初判表、現場圖、事故現場照片" Value="X" />
                            </asp:RadioButtonList>
                            <asp:Label ID="eHasCaseData_INS" runat="server" Text='<%# Eval("HasCaseData") %>' Visible="false" />
                            <br />
                            <asp:CheckBox ID="cbHasAccReport_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbHasAccReport_INS_CheckedChanged" Text="已申請行車事故鑑定" />
                            <asp:Label ID="eHasAccReport_INS" runat="server" Text='<%# Eval("HasAccReport") %>' Visible="false" />
                            <asp:Label ID="eERPCouseNo_INS" runat="server" Text='<%# Eval("ERPCouseNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動資料" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyMan_A_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_A") %>' Width="35%" />
                            <asp:Label ID="eModifyManName_A_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName_A") %>' Width="60%" />
                            <br />
                            <asp:Label ID="eModifyDate_A_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate_A","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                            <asp:Label ID="lbSubTitle_INS" runat="server" CssClass="titleText-Blue" Text="鑑定會審議結果" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbIsExemption_INS" runat="server" CssClass="text-Right-Blue" Text="駕駛責任" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:CheckBox ID="cbIsExemption_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnCheckedChanged="cbIsExemption_INS_CheckedChanged" Text="裁定免責" />
                            <asp:Label ID="eIsExemption_INS" runat="server" Text='<%# Eval("IsExemption") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPaidAmount_INS" runat="server" CssClass="text-Right-Blue" Text="已自付總額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePaidAmount_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PaidAmount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbInsuAmount_INS" runat="server" CssClass="text-Right-Blue" Text="保險理賠金" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eInsuAmount_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InsuAmount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPenaltyRatio_INS" runat="server" CssClass="text-Right-Blue" Text="罰款比例" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePenaltyRatio_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PenaltyRatio") %>' Width="75%" />
                            <a>％</a>
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPenalty_INS" runat="server" CssClass="text-Right-Blue" Text="罰款分擔額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePenalty_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Penalty") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                    </tr>
                </table>
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNewMain_Empty" runat="server" CssClass="button-Black" Text="新增" Width="120px" CommandName="New" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_Main" runat="server" CssClass="button-Black" Text="新增" CommandName="New" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit_Main" runat="server" CssClass="button-Blue" Text="修改" CommandName="Edit" Width="120px" />
                &nbsp;<asp:Button ID="bbDel_Main" runat="server" CssClass="button-Red" Text="刪除" OnClick="bbDel_Main_Click" Width="120px" />
                &nbsp;<asp:Button ID="bbPrintNote" runat="server" CssClass="button-Black" Text="列印審議通知" OnClick="bbPrintNote_Click" Width="120px" />
                &nbsp;<asp:Button ID="bbPrintReport" runat="server" CssClass="button-Blue" Text="列印肇事報告表" OnClick="bbPrintReport_Click" Width="120px" />
                &nbsp;<asp:Button ID="bbERPSync" runat="server" CssClass="button-Black" Text="ERP同步" OnClick="bbERPSync_Click" Width="120px" />
                &nbsp;<asp:Button ID="bbDelERP" runat="server" CausesValidation="false" CssClass="button-Red" Text="取消同步" OnClick="bbDelERP_Click" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCaseNo_A_List" runat="server" CssClass="text-Right-Blue" Text="肇事單號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eCaseNo_A_List" runat="server" Text='<%# Eval("CaseNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbHasInsurance_List" runat="server" CssClass="text-Right-Blue" Text="處理情況" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:CheckBox ID="cbHasInsurance_List" runat="server" CssClass="text-Left-Black" Text="出險" Enabled="false" Width="35%" />
                            <asp:CheckBox ID="cbCaseClose_List" runat="server" CssClass="text-Left-Black" Text="和解" Enabled="false" Width="35%" />
                            <asp:Label ID="eHasInsurance_List" runat="server" Text='<%# Eval("HasInsurance") %>' Visible="false" />
                            <asp:Label ID="eCaseClose_List" runat="server" Text='<%# Eval("CaseClose") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuildDate_List" runat="server" CssClass="text-Right-Blue" Text="出險日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBuildDate_List" runat="server" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBuildMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eBuildMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="35%" />
                            <asp:Label ID="eBuildManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildManName") %>' Width="60%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbReportDate_List" runat="server" CssClass="text-Right-Blue" Text="報告日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eReportDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ReportDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbInsuMan_List" runat="server" CssClass="text-Right-Blue" Text="保險經辦" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eInsuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InsuMan") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAnecdotalResRatio_List" runat="server" CssClass="text-Right-Blue" Text="肇責比例" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAnecdotalResRatio_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AnecdotalResRatio") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:CheckBox ID="cbIsNoDeduction_List" runat="server" CssClass="text-Left-Black" Text="免扣精勤" Enabled="false" />
                            <asp:Label ID="eIsNoDeduction_List" runat="server" Text='<%# Eval("IsNoDeduction") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDeductionDate_List" runat="server" CssClass="text-Right-Blue" Text="減發精勤日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eDeductionDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DeductionDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRemark_A_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                            <asp:TextBox ID="eRemark_A_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Remark_A") %>' TextMode="MultiLine" Enabled="false" Width="97%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                            <asp:Label ID="lbSubTitle_Driver_List" runat="server" CssClass="titleText-Blue" Text="駕駛員資料" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCar_ID_List" runat="server" CssClass="text-Right-Blue" Text="牌照號碼" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eCar_ID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDepName_List" runat="server" CssClass="text-Right-Blue" Text="所屬單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="35%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDriver_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eDriver_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver") %>' Width="35%" />
                            <asp:Label ID="eDriverName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="60%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbIDCardNo_List" runat="server" CssClass="text-Right-Blue" Text="身分證字號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eIDCardNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("IDCardNo") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbBirthday_List" runat="server" CssClass="text-Right-Blue" Text="出生日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eBirthday_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Birthday","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAssumeday_List" runat="server" CssClass="text-Right-Blue" Text="到職日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eAssumeday_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Assumeday","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbTelephoneNo_List" runat="server" CssClass="text-Right-Blue" Text="連絡電話" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eTelephoneNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TelephoneNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbAddress_List" runat="server" CssClass="text-Right-Blue" Text="地址" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                            <asp:Label ID="eAddress_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Address") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPersonDamage_List" runat="server" CssClass="text-Right-Blue" Text="體傷" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="4">
                            <asp:Label ID="ePersonDamage_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PersonDamage") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCarDamage_List" runat="server" CssClass="text-Right-Blue" Text="車 (財) 損" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="4">
                            <asp:Label ID="eCarDamage_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarDamage") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                            <asp:Label ID="lbSubTitle_CaseNote_List" runat="server" CssClass="titleText-Blue" Text="事故相關訊息" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCaseDate_List" runat="server" CssClass="text-Right-Blue" Text="事故日期 / 時間" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eCaseDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDate","{0:yyyy/MM/dd}") %>' Width="55%" />
                            <asp:Label ID="eCaseTime_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseTime") %>' Width="35%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbOutReportNo_List" runat="server" CssClass="text-Right-Blue" Text="登記聯單編號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eOutReportNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OutReportNo") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCasePosition_List" runat="server" CssClass="text-Right-Blue" Text="事故地點" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="4">
                            <asp:Label ID="eCasePosition_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CasePosition") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPoliceUnit_List" runat="server" CssClass="text-Right-Blue" Text="警方受理單位" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                            <asp:Label ID="ePoliceUnit_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PoliceUnit") %>' Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPoliceName_List" runat="server" CssClass="text-Right-Blue" Text="承辦員警" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePoliceName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PoliceName") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:CheckBox ID="cbHasVideo_List" runat="server" CssClass="text-Left-Black" Text="有行車影像" Enabled="false" Width="95%" />
                            <asp:Label ID="eHasVideo_List" runat="server" Text='<%# Eval("HasVideo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbNoVodeoReason_List" runat="server" CssClass="text-Right-Blue" Text="無影像原因" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eNoVideoReason_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("NoVideoReason") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCaseOccurrence_List" runat="server" CssClass="text-Right-Blue" Text="肇事經過" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                            <asp:TextBox ID="eCaseOccurrence_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseOccurrence") %>' Enabled="false" Width="97%" Height="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbHasCaseData_List" runat="server" CssClass="text-Right-Blue" Text="已申請資料" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="6">
                            <asp:RadioButtonList ID="rbHasCaseData_List" runat="server" CssClass="text-Left-Black" RepeatColumns="1" Enabled="false">
                                <asp:ListItem Text="有" Value="V" Selected="True" />
                                <asp:ListItem Text="無申請初判表、現場圖、事故現場照片" Value="X" />
                            </asp:RadioButtonList>
                            <asp:Label ID="eHasCaseData_List" runat="server" Text='<%# Eval("HasCaseData") %>' Visible="false" />
                            <br />
                            <asp:CheckBox ID="cbHasAccReport_List" runat="server" CssClass="text-Left-Black" Text="已申請行車事故鑑定" Enabled="false" />
                            <asp:Label ID="eHasAccReport_List" runat="server" Text='<%# Eval("HasAccReport") %>' Visible="false" />
                            <asp:Label ID="eERPCouseNo_List" runat="server" Text='<%# Eval("ERPCouseNo") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動資料" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:Label ID="eModifyMan_A_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_A") %>' Width="35%" />
                            <asp:Label ID="eModifyManName_A_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName_A") %>' Width="60%" />
                            <br />
                            <asp:Label ID="eModifyDate_A_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate_A","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                            <asp:Label ID="lbSubTitle_List" runat="server" CssClass="titleText-Blue" Text="鑑定會審議結果" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbIsExemption_List" runat="server" CssClass="text-Right-Blue" Text="駕駛責任" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:CheckBox ID="cbIsExemption_List" runat="server" CssClass="text-Left-Black" Text="裁定免責" Enabled="false" />
                            <asp:Label ID="eIsExemption_List" runat="server" Text='<%# Eval("IsExemption") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPaidAmount_List" runat="server" CssClass="text-Right-Blue" Text="已自付總額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePaidAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PaidAmount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbInsuAmount_List" runat="server" CssClass="text-Right-Blue" Text="保險理賠金" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eInsuAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("InsuAmount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPenaltyRatio_List" runat="server" CssClass="text-Right-Blue" Text="罰款比例" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePenaltyRatio_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PenaltyRatio") %>' Width="75%" />
                            <a>％</a>
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPenalty_List" runat="server" CssClass="text-Right-Blue" Text="罰款分擔額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePenalty_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Penalty") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                        <td class="ColHeight ColWidth-10Col" />
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsAnecdoteCaseA_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        SelectCommand="SELECT a.CaseNo, a.HasInsurance, a.CaseClose, a.DepName, a.BuildDate, b.[NAME] as BuildManName, a.Car_ID, a.DriverName, a.InsuMan, a.AnecdotalResRatio, a.IsNoDeduction, a.DeductionDate, 
       a.Remark
  FROM AnecdoteCase AS a left join Employee b on b.EmpNo = a.BuildMan WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsAnecdoteCaseA_Data" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select a.CaseNo, a.HasInsurance, a.DepNo, a.DepName, convert(varchar(10), a.BuildDate, 111) as BuildDate, a.BuildMan, 
       e.[Name] as BuildManName, a.Car_ID, a.Driver, a.DriverName, a.InsuMan, a.AnecdotalResRatio, a.IsNoDeduction, 
       convert(varchar(10), a.DeductionDate, 111) as DeductionDate, a.Remark as Remark_A, a.CaseOccurrence, a.ERPCouseNo, a.CaseClose, 
       a.IsExemption, a.PaidAmount, a.Penalty, a.PenaltyRatio, a.InsuAmount, a.IDCardNo, convert(varchar(10), a.Birthday, 111) as Birthday, 
       convert(varchar(10), a.Assumeday, 111) as Assumeday, a.TelephoneNo, a.[Address],a.PersonDamage, a.CarDamage,  
       convert(varchar(10), a.ReportDate, 111) as ReportDate, convert(varchar(10), a.CaseDate, 111) as CaseDate, 
       a.CaseTime, a.OutReportNo, a.CasePosition, a.PoliceUnit, a.PoliceName, a.HasVideo, a.NoVideoReason, a.HasCaseData, 
       a.HasAccReport, a.ModifyMan as ModifyMan_A, (select [Name] from Employee where EmpNo = a.ModifyMan) as ModifyManName_A, convert(varchar(10), a.ModifyDate, 111) as ModifyDate_A 
  from AnecdoteCase a left join Employee e on e.EmpNo = a.BuildMan 
 where a.CaseNo = @CaseNo">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridAnecdoteCaseA_List" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:Panel ID="plDetailDataShow" runat="server" CssClass="ShowPanel-Detail">
        <asp:GridView ID="gridAnecdoteCaseB_List" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px" CellPadding="4" DataKeyNames="CaseNoItems" DataSourceID="sdsAnecdoteCaseB_List" PageSize="5">
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
                <asp:BoundField DataField="DriverSharing" HeaderText="駕駛員負擔" SortExpression="DriverSharing" />
                <asp:BoundField DataField="CompanySharing" HeaderText="公司負擔" SortExpression="CompanySharing" />
                <asp:BoundField DataField="CarDamageAMT" HeaderText="車損金額" SortExpression="CarDamageAMT" />
                <asp:BoundField DataField="PersonDamageAMT" HeaderText="體傷金額" SortExpression="PersonDamageAMT" />
                <asp:BoundField DataField="RelationComp" HeaderText="對方賠付金額" SortExpression="RelationComp" />
                <asp:BoundField DataField="ReconciliationDate" DataFormatString="{0:d}" HeaderText="和解日期" SortExpression="ReconciliationDate" />
                <asp:BoundField DataField="PassengerInsu" HeaderText="PassengerInsu" SortExpression="PassengerInsu" Visible="False" />
                <asp:BoundField DataField="Remark" HeaderText="Remark" SortExpression="Remark" Visible="False" />
            </Columns>
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
        <div>
            <asp:Label ID="eErrorMSG_B" runat="server" CssClass="errorMessageText" Visible="false" />
        </div>
        <asp:FormView ID="fvAnecdoteCaseB_Data" runat="server" Width="100%" AllowPaging="True" DataKeyNames="CaseNoItems" DataSourceID="sdsAnecdoteCaseB_Data" OnDataBound="fvAnecdoteCaseB_Data_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbDetailOK_Edit" runat="server" CausesValidation="True" CssClass="button-Black" OnClick="bbDetailOK_Edit_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbDetailCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCaseNoItems_Edit" runat="server" CssClass="text-Right-Blue" Text="明細項次" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItems_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eCaseNo_B_Edit" runat="server" Text='<%# Eval("CaseNo") %>' Visible="false" />
                            <asp:Label ID="eCaseNoItems_Edit" runat="server" Text='<%# Eval("CaseNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelationship_Edit" runat="server" CssClass="text-Right-Blue" Text="對方姓名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eRelationship_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Relationship") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelGender_Edit" runat="server" CssClass="text-Right-Blue" Text="性別" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eRelGender_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelGender") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelTelNo1_Edit" runat="server" CssClass="text-Right-Blue" Text="連絡電話一" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eRelTelNo1_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelTelNo1") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelTelNo2_Edit" runat="server" CssClass="text-Right-Blue" Text="連絡電話二" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eRelTelNo2_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelTelNo2") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-10Col" colspan="5" />
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelCarType_Edit" runat="server" CssClass="text-Right-Blue" Text="車種" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:DropDownList ID="ddlRelCarType_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlRelCarType_Edit_SelectedIndexChanged" Width="95%">
                                <asp:ListItem Text="" Value="00" />
                                <asp:ListItem Text="行人 (含輪椅)" Value="01" />
                                <asp:ListItem Text="腳踏車" Value="02" />
                                <asp:ListItem Text="機車 (含電動二輪)" Value="03" />
                                <asp:ListItem Text="自小客貨" Value="04" />
                                <asp:ListItem Text="營小客貨" Value="05" />
                                <asp:ListItem Text="大客車" Value="06" />
                                <asp:ListItem Text="其他" Value="99" />
                            </asp:DropDownList>
                            <asp:Label ID="eRelCarType_Edit" runat="server" Text='<%# Eval("RelCarType") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelCar_ID_Edit" runat="server" CssClass="text-Right-Blue" Text="對方車號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eRelCar_ID_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelCar_ID") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_High ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelPersonDamage_Edit" runat="server" CssClass="text-Right-Blue" Text="體傷" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="4">
                            <asp:TextBox ID="eRelPersonDamage_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RelPersonDamage") %>' Height="97%" Width="97%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelCarDamage_Edit" runat="server" CssClass="text-Right-Blue" Text="車 (財) 損" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="4">
                            <asp:TextBox ID="eRelCarDamage_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RelCarDamage") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelNote_Edit" runat="server" CssClass="text-Right-Blue" Text="處理情形概述" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="9">
                            <asp:TextBox ID="eRelNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RelNote") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbEstimatedAmount_Edit" runat="server" CssClass="text-Right-Blue" Text="預估金額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eEstimatedAmount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EstimatedAmount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbThirdInsurance_Edit" runat="server" CssClass="text-Right-Blue" Text="第三責任險" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eThirdInsurance_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ThirdInsurance") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCompInsurance_Edit" runat="server" CssClass="text-Right-Blue" Text="強制險" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eCompInsurance_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompInsurance") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPassengerInsu_Edit" runat="server" CssClass="text-Right-Blue" Text="乘客險" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePassengerInsu_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PassengerInsu") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbReconciliationDate_Edit" runat="server" CssClass="text-Right-Blue" Text="和解日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eReconciliationDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ReconciliationDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCarDamageAMT_Edit" runat="server" CssClass="text-Right-Blue" Text="車損金額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eCarDamageAMT_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarDamageAMT") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPersonDamageAMT_Edit" runat="server" CssClass="text-Right-Blue" Text="體傷金額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePersonDamageAMT_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PersonDamageAMT") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCompanySharing_Edit" runat="server" CssClass="text-Right-Blue" Text="公司負擔" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eCompanySharing" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompanySharing") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDriverSharing_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛員負擔" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eDriverSharing_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverSharing") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelationComp_Edit" runat="server" CssClass="text-Right-Blue" Text="對方賠付" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eRelationComp_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelationComp") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_High ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRemark_B_Edit" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="7">
                            <asp:TextBox ID="eRemark_B_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Height="97%" Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="95%" />
                            <br />
                            <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                    </tr>
                </table>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbDetailOK_INS" runat="server" CausesValidation="True" CssClass="button-Black" OnClick="bbDetailOK_INS_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbDetailCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCaseNoItems_INS" runat="server" CssClass="text-Right-Blue" Text="明細項次" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItems_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eCaseNo_B_INS" runat="server" Text='<%# Eval("CaseNo") %>' Visible="false" />
                            <asp:Label ID="eCaseNoItems_INS" runat="server" Text='<%# Eval("CaseNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelationship_INS" runat="server" CssClass="text-Right-Blue" Text="對方姓名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eRelationship_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Relationship") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelGender_INS" runat="server" CssClass="text-Right-Blue" Text="性別" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eRelGender_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelGender") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelTelNo1_INS" runat="server" CssClass="text-Right-Blue" Text="連絡電話一" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eRelTelNo1_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelTelNo1") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelTelNo2_INS" runat="server" CssClass="text-Right-Blue" Text="連絡電話二" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eRelTelNo2_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelTelNo2") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-10Col" colspan="5" />
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelCarType_INS" runat="server" CssClass="text-Right-Blue" Text="車種" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:DropDownList ID="ddlRelCarType_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlRelCarType_INS_SelectedIndexChanged" Width="95%">
                                <asp:ListItem Text="" Value="00" />
                                <asp:ListItem Text="行人 (含輪椅)" Value="01" />
                                <asp:ListItem Text="腳踏車" Value="02" />
                                <asp:ListItem Text="機車 (含電動二輪)" Value="03" />
                                <asp:ListItem Text="自小客貨" Value="04" />
                                <asp:ListItem Text="營小客貨" Value="05" />
                                <asp:ListItem Text="大客車" Value="06" />
                                <asp:ListItem Text="其他" Value="99" />
                            </asp:DropDownList>
                            <asp:Label ID="eRelCarType_INS" runat="server" Text='<%# Eval("RelCarType") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelCar_ID_INS" runat="server" CssClass="text-Right-Blue" Text="對方車號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eRelCar_ID_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelCar_ID") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_High ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelPersonDamage_INS" runat="server" CssClass="text-Right-Blue" Text="體傷" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="4">
                            <asp:TextBox ID="eRelPersonDamage_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RelPersonDamage") %>' Height="97%" Width="97%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelCarDamage_INS" runat="server" CssClass="text-Right-Blue" Text="車 (財) 損" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="4">
                            <asp:TextBox ID="eRelCarDamage_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RelCarDamage") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelNote_INS" runat="server" CssClass="text-Right-Blue" Text="處理情形概述" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="9">
                            <asp:TextBox ID="eRelNote_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("RelNote") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbEstimatedAmount_INS" runat="server" CssClass="text-Right-Blue" Text="預估金額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eEstimatedAmount_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EstimatedAmount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbThirdInsurance_INS" runat="server" CssClass="text-Right-Blue" Text="第三責任險" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eThirdInsurance_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ThirdInsurance") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCompInsurance_INS" runat="server" CssClass="text-Right-Blue" Text="強制險" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eCompInsurance_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompInsurance") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPassengerInsu_INS" runat="server" CssClass="text-Right-Blue" Text="乘客險" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePassengerInsu_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PassengerInsu") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbReconciliationDate_INS" runat="server" CssClass="text-Right-Blue" Text="和解日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eReconciliationDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ReconciliationDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCarDamageAMT_INS" runat="server" CssClass="text-Right-Blue" Text="車損金額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eCarDamageAMT_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarDamageAMT") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPersonDamageAMT_INS" runat="server" CssClass="text-Right-Blue" Text="體傷金額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="ePersonDamageAMT_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PersonDamageAMT") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCompanySharing_INS" runat="server" CssClass="text-Right-Blue" Text="公司負擔" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eCompanySharing" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompanySharing") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDriverSharing_INS" runat="server" CssClass="text-Right-Blue" Text="駕駛員負擔" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eDriverSharing_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverSharing") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelationComp_INS" runat="server" CssClass="text-Right-Blue" Text="對方賠付" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:TextBox ID="eRelationComp_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelationComp") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_High ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRemark_B_INS" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="7">
                            <asp:TextBox ID="eRemark_B_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Height="97%" Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="95%" />
                            <br />
                            <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                    </tr>
                </table>
            </InsertItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNewDetail_Empty" runat="server" CausesValidation="false" CssClass="button-Black" CommandName="New" Text="新增明細" Width="120px" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew_Detail" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增明細" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit_Detail" runat="server" CausesValidation="False" CssClass="button-Blue" CommandName="Edit" Text="修改" Width="120px" />
                &nbsp;<asp:Button ID="bbDel_Detail" runat="server" CausesValidation="False" CssClass="button-Red" OnClick="bbDel_Detail_Click" Text="刪除" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCaseNoItems_List" runat="server" CssClass="text-Right-Blue" Text="明細項次" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eItems_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Items") %>' Width="95%" />
                            <asp:Label ID="eCaseNo_B_List" runat="server" Text='<%# Eval("CaseNo") %>' Visible="false" />
                            <asp:Label ID="eCaseNoItems_List" runat="server" Text='<%# Eval("CaseNoItems") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelationship_List" runat="server" CssClass="text-Right-Blue" Text="對方姓名" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eRelationship_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Relationship") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelGender_List" runat="server" CssClass="text-Right-Blue" Text="性別" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eRelGender_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelGender") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelTelNo1_List" runat="server" CssClass="text-Right-Blue" Text="連絡電話一" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eRelTelNo1_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelTelNo1") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelTelNo2_List" runat="server" CssClass="text-Right-Blue" Text="連絡電話二" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eRelTelNo2_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelTelNo2") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-10Col" colspan="5" />
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelCarType_List" runat="server" CssClass="text-Right-Blue" Text="車種" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                            <asp:DropDownList ID="ddlRelCarType_List" runat="server" CssClass="text-Left-Black" Enabled="false" Width="95%">
                                <asp:ListItem Text="" Value="00" />
                                <asp:ListItem Text="行人 (含輪椅)" Value="01" />
                                <asp:ListItem Text="腳踏車" Value="02" />
                                <asp:ListItem Text="機車 (含電動二輪)" Value="03" />
                                <asp:ListItem Text="自小客貨" Value="04" />
                                <asp:ListItem Text="營小客貨" Value="05" />
                                <asp:ListItem Text="大客車" Value="06" />
                                <asp:ListItem Text="其他" Value="99" />
                            </asp:DropDownList>
                            <asp:Label ID="eRelCarType_List" runat="server" Text='<%# Eval("RelCarType") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelCar_ID_List" runat="server" CssClass="text-Right-Blue" Text="對方車號" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eRelCar_ID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelCar_ID") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_High ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelPersonDamage_List" runat="server" CssClass="text-Right-Blue" Text="體傷" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="4">
                            <asp:TextBox ID="eRelPersonDamage_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("RelPersonDamage") %>' Height="97%" Width="97%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelCarDamage_List" runat="server" CssClass="text-Right-Blue" Text="車 (財) 損" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="4">
                            <asp:TextBox ID="eRelCarDamage_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("RelCarDamage") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelNote_List" runat="server" CssClass="text-Right-Blue" Text="處理情形概述" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col" colspan="9">
                            <asp:TextBox ID="eRelNote_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("RelNote") %>' Height="97%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbEstimatedAmount_List" runat="server" CssClass="text-Right-Blue" Text="預估金額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eEstimatedAmount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EstimatedAmount") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbThirdInsurance_List" runat="server" CssClass="text-Right-Blue" Text="第三責任險" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eThirdInsurance_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ThirdInsurance") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCompInsurance_List" runat="server" CssClass="text-Right-Blue" Text="強制險" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eCompInsurance_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompInsurance") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPassengerInsu_List" runat="server" CssClass="text-Right-Blue" Text="乘客險" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePassengerInsu_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PassengerInsu") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbReconciliationDate_List" runat="server" CssClass="text-Right-Blue" Text="和解日期" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eReconciliationDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ReconciliationDate","{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCarDamageAMT_List" runat="server" CssClass="text-Right-Blue" Text="車損金額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eCarDamageAMT_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CarDamageAMT") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbPersonDamageAMT_List" runat="server" CssClass="text-Right-Blue" Text="體傷金額" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="ePersonDamageAMT_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("PersonDamageAMT") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbCompanySharing_List" runat="server" CssClass="text-Right-Blue" Text="公司負擔" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eCompanySharing" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CompanySharing") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbDriverSharing_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員負擔" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eDriverSharing_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverSharing") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRelationComp_List" runat="server" CssClass="text-Right-Blue" Text="對方賠付" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="eRelationComp_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RelationComp") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_High ColBorder ColWidth-10Col">
                            <asp:Label ID="lbRemark_B_List" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        </td>
                        <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="7">
                            <asp:TextBox ID="eRemark_B_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Height="97%" Width="97%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-10Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動人" Width="95%" />
                        </td>
                        <td class="MultiLine_High ColBorder ColWidth-10Col">
                            <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="95%" />
                            <br />
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                        <td class="ColWidth-10Col" />
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsAnecdoteCaseB_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, PassengerInsu, Remark FROM AnecdoteCaseB WHERE (CaseNo = @CaseNo)">
        <SelectParameters>
            <asp:ControlParameter ControlID="fvAnecdoteCaseA_Data" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsAnecdoteCaseB_Data" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, PassengerInsu, Remark, RelGender, RelTelNo1, RelTelNo2, RelCarType, RelPersonDamage, RelCarDamage, RelNote, ModifyMan, ModifyDate FROM AnecdoteCaseB
where CaseNoItems = @CaseNoItems">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridAnecdoteCaseB_List" Name="CaseNoItems" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel">
        <asp:Button ID="bbCloseReport" runat="server" CssClass="button-Red" Text="結束預覽" OnClick="bbCloseReport_Click" Width="120px" />
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
</asp:Content>
