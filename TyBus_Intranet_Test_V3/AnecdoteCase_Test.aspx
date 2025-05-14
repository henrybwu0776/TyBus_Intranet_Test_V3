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
        <asp:GridView ID="gridAnecdoteCaseA_List" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            CellPadding="4" ForeColor="#333333" DataKeyNames="CaseNo" GridLines="None" PageSize="5" Width="100%" OnPageIndexChanging="gridAnecdoteCaseA_List_PageIndexChanging" OnSelectedIndexChanged="gridAnecdoteCaseA_List_SelectedIndexChanged">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
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
        <div>
            <asp:Button ID="bbNew_Main" runat="server" CssClass="button-Black" Text="新增" OnClick="bbNew_Main_Click" Width="120px" />
            <asp:Button ID="bbEdit_Main" runat="server" CssClass="button-Blue" Text="修改" OnClick="bbEdit_Main_Click" Width="120px" />
            <asp:Button ID="bbDel_Main" runat="server" CssClass="button-Red" Text="刪除" OnClick="bbDel_Main_Click" Width="120px" />
            <asp:Button ID="bbPrintNote" runat="server" CssClass="button-Black" Text="列印審議通知" OnClick="bbPrintNote_Click" />
            <asp:Button ID="bbPrintReport" runat="server" CssClass="button-Blue" Text="列印肇事報告表" OnClick="bbPrintReport_Click" />
            <asp:Button ID="bbERPSync" runat="server" CssClass="button-Black" Text="ERP同步" OnClick="bbERPSync_Click" Width="120px" />
            <asp:Button ID="bbOK_Main" runat="server" CssClass="button-Blue" Text="確定" OnClick="bbOK_Main_Click" Width="120px" />
            <asp:Button ID="bbCancel_Main" runat="server" CssClass="button-Red" Text="取消" OnClick="bbCancel_Main_Click" Width="120px" />
        </div>
        <asp:Panel ID="plMainData" runat="server" CssClass="ShowPanel">
            <div>
                <asp:Label ID="eErrorMSG_A" runat="server" CssClass="errorMessageText" Visible="false" />
            </div>
            <table class="TableSetting">
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbCaseNo_A" runat="server" CssClass="text-Right-Blue" Text="肇事單號" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eCaseNo_A" runat="server" CssClass="text-Left-Black" Text='<%# vdCaseNo_A %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbHasInsurance" runat="server" CssClass="text-Right-Blue" Text="處理情況" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                        <asp:CheckBox ID="cbHasInsurance" runat="server" CssClass="text-Left-Black" Text="出險" AutoPostBack="true" OnCheckedChanged="cbHasInsurance_CheckedChanged" Width="35%" />
                        <asp:CheckBox ID="cbCaseClose" runat="server" CssClass="text-Left-Black" Text="和解" AutoPostBack="true" OnCheckedChanged="cbCaseClose_CheckedChanged" Width="35%" />
                        <asp:Label ID="eHasInsurance" runat="server" Text='<%# vdHasInsurance %>' Visible="false" />
                        <asp:Label ID="eCaseClose" runat="server" Text='<%# vdCaseClose %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbBuildDate" runat="server" CssClass="text-Right-Blue" Text="出險日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eBuildDate" runat="server" CssClass="text-Left-Black" Text='<%# vdBuildDate %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbBuildMan" runat="server" CssClass="text-Right-Blue" Text="建檔人" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                        <asp:TextBox ID="eBuildMan" runat="server" CssClass="text-Left-Black" Text='<%# vdBuildMan %>' AutoPostBack="true" OnTextChanged="eBuildMan_TextChanged" Width="35%" />
                        <asp:Label ID="eBuildManName" runat="server" CssClass="text-Left-Black" Text='<%# vdBuildManName %>' Width="60%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbReportDate" runat="server" CssClass="text-Right-Blue" Text="報告日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eReportDate" runat="server" CssClass="text-Left-Black" Text='<%# vdReportDate %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbInsuMan" runat="server" CssClass="text-Right-Blue" Text="保險經辦" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eInsuMan" runat="server" CssClass="text-Left-Black" Text='<%# vdInsuMan %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbAnecdotalResRatio" runat="server" CssClass="text-Right-Blue" Text="肇責比例" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eAnecdotalResRatio" runat="server" CssClass="text-Left-Black" Text='<%# vdAnecdotalResRatio %>' Width="75%" />
                        <a>％</a>
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                        <asp:CheckBox ID="cbIsNoDeduction" runat="server" CssClass="text-Left-Black" Text="免扣精勤" AutoPostBack="true" OnCheckedChanged="cbIsNoDeduction_CheckedChanged" />
                        <asp:Label ID="eIsNoDeduction" runat="server" Text='<%# vdIsNoDeduction %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbDeductionDate" runat="server" CssClass="text-Right-Blue" Text="減發精勤日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eDeductionDate" runat="server" CssClass="text-Left-Black" Text='<%# vdDeductionDate %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbRemark" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                    </td>
                    <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                        <asp:TextBox ID="eRemark" runat="server" CssClass="text-Left-Black" Text='<%# vdRemark_A %>' TextMode="MultiLine" Width="97%" Height="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                        <asp:Label ID="lbSubTitle_Driver" runat="server" CssClass="titleText-Blue" Text="駕駛員資料" Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbCar_ID" runat="server" CssClass="text-Right-Blue" Text="牌照號碼" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eCar_ID" runat="server" CssClass="text-Left-Black" Text='<%# vdCar_ID %>' AutoPostBack="true" OnTextChanged="eCar_ID_TextChanged" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbDepName" runat="server" CssClass="text-Right-Blue" Text="所屬單位" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                        <asp:TextBox ID="eDepNo" runat="server" CssClass="text-Left-Black" Text='<%# vdDepNo %>' AutoPostBack="true" OnTextChanged="eDepNo_TextChanged" Width="35%" />
                        <asp:Label ID="eDepName" runat="server" CssClass="text-Left-Black" Text='<%# vdDepName %>' Width="60%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbDriver" runat="server" CssClass="text-Right-Blue" Text="駕駛員" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                        <asp:TextBox ID="eDriver" runat="server" CssClass="text-Left-Black" Text='<%# vdDriver %>' AutoPostBack="true" OnTextChanged="eDriver_TextChanged" Width="35%" />
                        <asp:Label ID="eDriverName" runat="server" CssClass="text-Left-Black" Text='<%# vdDriverName %>' Width="60%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbIDCardNo" runat="server" CssClass="text-Right-Blue" Text="身分證字號" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eIDCardNo" runat="server" CssClass="text-Left-Black" Text='<%# vdIDCardNo %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbBirthday" runat="server" CssClass="text-Right-Blue" Text="出生日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eBirthday" runat="server" CssClass="text-Left-Black" Text='<%# vdBirthday %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbAssumeday" runat="server" CssClass="text-Right-Blue" Text="到職日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eAssumeday" runat="server" CssClass="text-Left-Black" Text='<%# vdAssumeday %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbTelephoneNo" runat="server" CssClass="text-Right-Blue" Text="電話號碼" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eTelephoneNo" runat="server" CssClass="text-Left-Black" Text='<%# vdTelephoneNo %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbAddress" runat="server" CssClass="text-Right-Blue" Text="地址" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:TextBox ID="eAddress" runat="server" CssClass="text-Left-Black" Text='<%# vdAddress %>' Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbPersonDamage" runat="server" CssClass="text-Right-Blue" Text="體傷" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="4">
                        <asp:TextBox ID="ePersonDamage" runat="server" CssClass="text-Left-Black" Text='<%# vdPersonDamage %>' Width="97%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbCarDamage" runat="server" CssClass="text-Right-Blue" Text="車 (財) 損" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="4">
                        <asp:TextBox ID="eCarDamage" runat="server" CssClass="text-Left-Black" Text='<%# vdCarDamage %>' Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                        <asp:Label ID="lbSubTitle_CaseNote" runat="server" CssClass="titleText-Blue" Text="事故相關訊息" Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbCaseDate" runat="server" CssClass="text-Right-Blue" Text="事故日期 / 時間" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                        <asp:TextBox ID="eCaseDate" runat="server" CssClass="text-Left-Black" Text='<%# vdCaseDate %>' Width="55%" />
                        <asp:TextBox ID="eCaseTime" runat="server" CssClass="text-Left-Black" Text='<%# vdCaseTime %>' Width="35%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbOutReportNo" runat="server" CssClass="text-Right-Blue" Text="登記聯單編號" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eOutReportNo" runat="server" CssClass="text-Left-Black" Text='<%# vdOutReportNo %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbCasePosition" runat="server" CssClass="text-Right-Blue" Text="事故地點" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="4">
                        <asp:TextBox ID="eCasePosition" runat="server" CssClass="text-Left-Black" Text='<%# vdCasePosition %>' Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbPoliceUnit" runat="server" CssClass="text-Right-Blue" Text="警方受理單位" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3">
                        <asp:TextBox ID="ePoliceUnit" runat="server" CssClass="text-Left-Black" Text='<%# vdPoliceUnit %>' Width="97%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbPpliceName" runat="server" CssClass="text-Right-Blue" Text="承辦員警" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="ePoliceName" runat="server" CssClass="text-Left-Black" Text='<%# vdPoliceName %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:CheckBox ID="cbHasVideo" runat="server" CssClass="text-Left-Black" Text="有行車影像" AutoPostBack="true" OnCheckedChanged="cbHasVideo_CheckedChanged" Width="95%" />
                        <asp:Label ID="eHasVideo" runat="server" Text='<%# vdHasVideo %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbNoVodeoReason" runat="server" CssClass="text-Right-Blue" Text="無影像原因" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                        <asp:TextBox ID="eNoVideoReason" runat="server" CssClass="text-Left-Black" Text='<%# vdNoVideoReason %>' Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbCaseOccurrence" runat="server" CssClass="text-Right-Blue" Text="肇事經過" Width="95%" />
                    </td>
                    <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                        <asp:TextBox ID="eCaseOccurrence" runat="server" CssClass="text-Left-Black" Text='<%# vdCaseOccurrence %>' TextMode="MultiLine" Width="97%" Height="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbRemark_A" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                    </td>
                    <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                        <asp:TextBox ID="eRemark_A" runat="server" CssClass="text-Left-Black" Text='<%# vdRemark_A %>' TextMode="MultiLine" Width="97%" Height="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbHasCaseData" runat="server" CssClass="text-Right-Blue" Text="已申請資料" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="9">
                        <asp:RadioButtonList ID="rbHasCaseData" runat="server" CssClass="text-Left-Black" RepeatColumns="1" OnSelectedIndexChanged="rbHasCaseData_SelectedIndexChanged">
                            <asp:ListItem Text="有" Value="V" Selected="True" />
                            <asp:ListItem Text="無申請初判表、現場圖、事故現場照片" Value="X" />
                        </asp:RadioButtonList>
                        <asp:Label ID="eHasCaseData" runat="server" Text='<%# vdHasCaseData %>' Visible="false" />
                        <br />
                        <asp:CheckBox ID="cbHasAccReport" runat="server" CssClass="text-Left-Black" Text="已申請行車事故鑑定" AutoPostBack="true" OnCheckedChanged="cbHasAccReport_CheckedChanged" />
                        <asp:Label ID="eHasAccReport" runat="server" Text='<%# vdHasAccReport %>' Visible="false" />
                        <asp:Label ID="eERPCouseNo" runat="server" Text='<%# vdERPCouseNo %>' Visible="false" />
                        <asp:Label ID="eModifyMan_A" runat="server" Text='<%# vdModifyMan_A %>' Visible="false" />
                        <asp:Label ID="eModifyDate_A" runat="server" Text='<%# vdModifyDate_A %>' Visible="false" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                        <asp:Label ID="lbSubTitle" runat="server" CssClass="titleText-Blue" Text="鑑定會審議結果" Width="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbIsExemption" runat="server" CssClass="text-Right-Blue" Text="駕駛責任" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:CheckBox ID="cbIsExemption" runat="server" CssClass="text-Left-Black" Text="裁定免責" AutoPostBack="true" OnCheckedChanged="cbIsExemption_CheckedChanged" />
                        <asp:Label ID="eIsExemption" runat="server" Text='<%# vdIsExemption %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbPaidAmount" runat="server" CssClass="text-Right-Blue" Text="已自付總額" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="ePaidAmount" runat="server" CssClass="text-Left-Black" Text='<%# vdPaidAmount %>' Width="95%" OnTextChanged="ePaidAmount_TextChanged" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbInsuAmount" runat="server" CssClass="text-Right-Blue" Text="保險理賠金" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eInsuAmount" runat="server" CssClass="text-Left-Black" Text='<%# vdInsuAmount %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbPenaltyRatio" runat="server" CssClass="text-Right-Blue" Text="罰款比例" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="ePenaltyRatio" runat="server" CssClass="text-Left-Black" Text='<%# vdPenaltyRatio %>' Width="75%" />
                        <a>％</a>
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbPenalty" runat="server" CssClass="text-Right-Blue" Text="罰款分擔額" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="ePenalty" runat="server" CssClass="text-Left-Black" Text='<%# vdPenalty %>' Width="95%" />
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
        </asp:Panel>
    </asp:Panel>
    <div />
    <asp:Panel ID="plDetailDataShow" runat="server" CssClass="ShowPanel-Detail">
        <asp:GridView ID="gridAnecdoteCaseB_List" runat="server" AllowPaging="True" CellPadding="4" ForeColor="#333333" GridLines="None" PageSize="5" Width="100%" OnPageIndexChanging="gridAnecdoteCaseB_List_PageIndexChanging" OnSelectedIndexChanged="gridAnecdoteCaseB_List_SelectedIndexChanged" AutoGenerateColumns="False" DataKeyNames="CaseNoItems">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowCancelButton="False" ShowSelectButton="True" />
                <asp:BoundField DataField="CaseNo" HeaderText="CaseNo" Visible="False" />
                <asp:BoundField DataField="Items" HeaderText="項次" />
                <asp:BoundField DataField="CaseNoItems" HeaderText="CaseNoItems" Visible="False" />
                <asp:BoundField DataField="Relationship" HeaderText="對方姓名" />
                <asp:BoundField DataField="RelCarType" HeaderText="對方車種" />
                <asp:BoundField DataField="RelCar_ID" HeaderText="牌照號碼" />
                <asp:BoundField DataField="RelTelNo1" HeaderText="連絡電話" />
                <asp:BoundField DataField="ReconciliationDate" HeaderText="和解日期" />
            </Columns>
            <FooterStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#FFCC66" ForeColor="#333333" HorizontalAlign="Center" />
            <RowStyle BackColor="#FFFBD6" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="Navy" />
            <SortedAscendingCellStyle BackColor="#FDF5AC" />
            <SortedAscendingHeaderStyle BackColor="#4D0000" />
            <SortedDescendingCellStyle BackColor="#FCF6C0" />
            <SortedDescendingHeaderStyle BackColor="#820000" />
        </asp:GridView>
        <div>
            <asp:Button ID="bbNew_B" runat="server" CssClass="button-Black" Text="新增" Width="120px" OnClick="bbNew_B_Click" />
            <asp:Button ID="bbEdit_B" runat="server" CssClass="button-Blue" Text="修改" Width="120px" OnClick="bbEdit_B_Click" />
            <asp:Button ID="bbDel_B" runat="server" CssClass="button-Red" Text="刪除" Width="120px" OnClick="bbDel_B_Click" />
            <asp:Button ID="bbOK_B" runat="server" CssClass="button-Blue" Text="確定" Width="120px" OnClick="bbOK_B_Click" />
            <asp:Button ID="bbCancel_B" runat="server" CssClass="button-Red" Text="取消" Width="120px" OnClick="bbCancel_B_Click" />
        </div>
        <asp:Panel ID="plDetailData" runat="server" CssClass="ShowPanel-Detail">
            <div>
                <asp:Label ID="eErrorMSG_B" runat="server" CssClass="errorMessageText" Visible="false" />
            </div>
            <table class="TableSetting">
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbCaseNo_B" runat="server" CssClass="text-Right-Blue" Text="單號" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eCaseNo_B" runat="server" CssClass="text-Left-Black" Text='<%# vdCaseNo_B %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbCaseNoItems" runat="server" CssClass="text-Right-Blue" Text="明細項次" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="eItems" runat="server" CssClass="text-Left-Black" Text='<%# vdItems %>' Width="95%" />
                        <asp:Label ID="eCaseNoItems" runat="server" Text='<%# vdCaseNoItems %>' Visible="false" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbRelationship" runat="server" CssClass="text-Right-Blue" Text="對方姓名" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eRelationship" runat="server" CssClass="text-Left-Black" Text='<%# vdRelationship %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbRelGender" runat="server" CssClass="text-Right-Blue" Text="性別" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eRelGender" runat="server" CssClass="text-Left-Black" Text='<%# vdRelGender %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbRelTelNo1" runat="server" CssClass="text-Right-Blue" Text="連絡電話一" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eRelTelNo1" runat="server" CssClass="text-Left-Black" Text='<%# vdRelTelNo1 %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="3" />
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbRelCarType" runat="server" CssClass="text-Right-Blue" Text="車種" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col" colspan="2">
                        <asp:DropDownList ID="ddlRelCarType" runat="server" CssClass="text-Left-Black" Width="95%" AutoPostBack="true" OnSelectedIndexChanged="ddlRelCarType_SelectedIndexChanged">
                            <asp:ListItem Text="" Value="00" />
                            <asp:ListItem Text="行人 (含輪椅)" Value="01" />
                            <asp:ListItem Text="腳踏車" Value="02" />
                            <asp:ListItem Text="機車 (含電動二輪)" Value="03" />
                            <asp:ListItem Text="自小客貨" Value="04" />
                            <asp:ListItem Text="營小客貨" Value="05" />
                            <asp:ListItem Text="大客車" Value="06" />
                            <asp:ListItem Text="其他" Value="99" />
                        </asp:DropDownList>
                        <asp:Label ID="eRelCarType" runat="server" Text='<%# vdRelCarType %>' />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbRelCar_ID" runat="server" CssClass="text-Right-Blue" Text="對方車號" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eRelCar_ID" runat="server" CssClass="text-Left-Black" Text='<%# vdRelCar_ID %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbRelTelNo2" runat="server" CssClass="text-Right-Blue" Text="連絡電話二" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eRelTelNo2" runat="server" CssClass="text-Left-Black" Text='<%# vdRelTelNo2 %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColBorder ColWidth-10Col MultiLine_High">
                        <asp:Label ID="lbRelPersonDamage" runat="server" CssClass="text-Right-Blue" Text="體傷" Width="95%" />
                    </td>
                    <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="4">
                        <asp:TextBox ID="eRelPersonDamage" runat="server" CssClass="text-Left-Black" Text='<%# vdRelPersonDamage %>' Width="97%" Height="97%" />
                    </td>
                    <td class="ColBorder ColWidth-10Col MultiLine_High">
                        <asp:Label ID="lbRelCarDamage" runat="server" CssClass="text-Right-Blue" Text="車 (財) 損" Width="95%" />
                    </td>
                    <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="4">
                        <asp:TextBox ID="eRelCarDamage" runat="server" CssClass="text-Left-Black" Text='<%# vdRelCarDamage %>' Width="97%" Height="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbRelNote" runat="server" CssClass="text-Right-Blue" Text="處理情形概述" Width="95%" />
                    </td>
                    <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                        <asp:TextBox ID="eRelNote" runat="server" CssClass="text-Left-Black" Text='<%# vdRelNote %>' TextMode="MultiLine" Width="97%" Height="97%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbEstimatedAmount" runat="server" CssClass="text-Right-Blue" Text="預估金額" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eEstimatedAmount" runat="server" CssClass="text-Left-Black" Text='<%# vdEstimatedAmount %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbThirdInsurance" runat="server" CssClass="text-Right-Blue" Text="第三責任險" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eThirdInsurance" runat="server" CssClass="text-Left-Black" Text='<%# vdThirdInsurance %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbCompInsurance" runat="server" CssClass="text-Right-Blue" Text="強制險" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eCompInsurance" runat="server" CssClass="text-Left-Black" Text='<%# vdCompInsurance %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbPassengerInsu" runat="server" CssClass="text-Right-Blue" Text="乘客險" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="ePassengerInsu" runat="server" CssClass="text-Left-Black" Text='<%# vdPassengerInsu %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbReconciliationDate" runat="server" CssClass="text-Right-Blue" Text="和解日期" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eReconciliationDate" runat="server" CssClass="text-Left-Black" Text='<%# vdReconciliationDate %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbCarDamageAMT" runat="server" CssClass="text-Right-Blue" Text="車損金額" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eCarDamageAMT" runat="server" CssClass="text-Left-Black" Text='<%# vdCarDamageAMT %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbPersonDamageAMT" runat="server" CssClass="text-Right-Blue" Text="體傷金額" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="ePersonDamageAMT" runat="server" CssClass="text-Left-Black" Text='<%# vdPersonDamageAMT %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbCompanySharing" runat="server" CssClass="text-Right-Blue" Text="公司負擔" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eCompanySharing" runat="server" CssClass="text-Left-Black" Text='<%# vdCompanySharing %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbDriverSharing" runat="server" CssClass="text-Right-Blue" Text="駕駛員負擔" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eDriverSharing" runat="server" CssClass="text-Left-Black" Text='<%# vdDriverSharing %>' Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbRelationComp" runat="server" CssClass="text-Right-Blue" Text="對方賠付" Width="95%" />
                    </td>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:TextBox ID="eRelationComp" runat="server" CssClass="text-Left-Black" Text='<%# vdRelationComp %>' Width="95%" />
                    </td>
                </tr>
                <tr>
                    <td class="ColHeight ColBorder ColWidth-10Col">
                        <asp:Label ID="lbRemark_B" runat="server" CssClass="text-Right-Blue" Text="備註" Width="95%" />
                        <asp:Label ID="eModifyMan_B" runat="server" Text='<%# vdModifyMan_B %>' Visible="false" />
                        <asp:Label ID="eModifyDate_B" runat="server" Text='<%# vdModifyDate_B %>' Visible="false" />
                    </td>
                    <td class="ColBorder ColWidth-10Col MultiLine_High" colspan="9">
                        <asp:TextBox ID="eRemark_B" runat="server" CssClass="text-Left-Black" Text='<%# vdRemark_B %>' Width="97%" Height="97%" TextMode="MultiLine" />
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
        </asp:Panel>
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
    <asp:SqlDataSource ID="dsAnecdoteA" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="SELECT CaseNo, 
       case when HasInsurance = 1 then '是' else '否' end HasInsurance, 
       DepNo, DepName, BuildDate, BuildMan, 
       (SELECT NAME FROM Employee WHERE (EMPNO = AnecdoteCase.BuildMan)) AS BuildManName, 
       Car_ID, Driver, DriverName, InsuMan, AnecdotalResRatio, 
       case when IsNoDeduction = 1 then '是' else '否' end IsNoDeduction, 
       DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose, 
       IsExemption, PaidAmount, InsuAmount, Penalty, PenaltyRatio 
  FROM AnecdoteCase 
 where isnull(CaseNo, '') = ''"></asp:SqlDataSource>
</asp:Content>
