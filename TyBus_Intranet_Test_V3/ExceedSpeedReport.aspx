<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="ExceedSpeedReport.aspx.cs" Inherits="TyBus_Intranet_Test_V3.ExceedSpeedReport" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="ExceedSpeedReportForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">各單位遊覽車超速檢查表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCaseDate_Search" runat="server" CssClass="text-Right-Blue" Text="檢查日：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:TextBox ID="eCaseDateYear_Search" runat="server" AutoPostBack="true" OnTextChanged="ddlMonthStep_Search_SelectedIndexChanged" CssClass="text-Left-Black" Width="10%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:TextBox ID="eCaseDateMonth_Search" runat="server" AutoPostBack="true" OnTextChanged="ddlMonthStep_Search_SelectedIndexChanged" CssClass="text-Left-Black" Width="10%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text=" 月 " />
                    <asp:DropDownList ID="ddlMonthStep_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlMonthStep_Search_SelectedIndexChanged" Width="10%">
                        <asp:ListItem Value="10" Text="上旬" Selected="True" />
                        <asp:ListItem Value="11" Text="中旬" />
                        <asp:ListItem Value="12" Text="下旬" />
                    </asp:DropDownList>
                    <asp:Label ID="eCaseDateS_Search" runat="server" CssClass="text-Left-Black" Width="20%" />
                    <asp:Label ID="lbSplit_3" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:Label ID="eCaseDateE_Search" runat="server" CssClass="text-Left-Black" Width="20%" />
                    <asp:Label ID="eMonthStep_Search" runat="server" Visible="false" Width="5%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbCreate_Search" runat="server" CssClass="button-Red" OnClick="bbCreate_Search_Click" Text="產生旬記錄(空白)" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" OnClick="bbSearch_Click" Text="查詢" Width="90%" />
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
        <asp:Button ID="bbPrint" runat="server" CssClass="button-Black" OnClick="bbPrint_Click" Text="預覽報表" Width="120px" />
        <asp:Button ID="bbExcel" runat="server" CssClass="button-Blue" OnClick="bbExcel_Click" Text="匯出 EXCEL" Width="120px" />
        <asp:GridView ID="gridExceedSpeedReport_List" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px" CellPadding="4" DataKeyNames="CaseNo" DataSourceID="sdsExceedSpeedReport_List" Width="100%">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="CaseNo" HeaderText="CaseNo" ReadOnly="True" SortExpression="CaseNo" Visible="False" />
                <asp:BoundField DataField="CaseYM" HeaderText="CaseYM" SortExpression="CaseYM" Visible="False" />
                <asp:BoundField DataField="MonthStep" HeaderText="MonthStep" SortExpression="MonthStep" Visible="False" />
                <asp:BoundField DataField="CaseDateS" DataFormatString="{0:D}" HeaderText="開始日" SortExpression="CaseDateS" />
                <asp:BoundField DataField="CaseDateE" DataFormatString="{0:D}" HeaderText="結束日" SortExpression="CaseDateE" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="部門" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="Car_ID" HeaderText="車號" SortExpression="Car_ID" />
                <asp:BoundField DataField="Driver" HeaderText="Driver" SortExpression="Driver" Visible="False" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員" SortExpression="DriverName" />
                <asp:BoundField DataField="Exception" HeaderText="異常項目" SortExpression="Exception" />
                <asp:BoundField DataField="AbnormalValue" HeaderText="異常數值" SortExpression="AbnormalValue" />
                <asp:BoundField DataField="Attachment" HeaderText="附件編號" SortExpression="Attachment" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" />
                <asp:BoundField DataField="Inspector" HeaderText="Inspector" SortExpression="Inspector" Visible="False" />
                <asp:BoundField DataField="InspectorName" HeaderText="查核人" SortExpression="InspectorName" />
                <asp:BoundField DataField="BuDate" HeaderText="BuDate" SortExpression="BuDate" Visible="False" />
                <asp:BoundField DataField="BuMan" HeaderText="BuMan" SortExpression="BuMan" Visible="False" />
                <asp:BoundField DataField="BuManName" HeaderText="BuManName" ReadOnly="True" SortExpression="BuManName" Visible="False" />
                <asp:BoundField DataField="ModifyDate" HeaderText="ModifyDate" SortExpression="ModifyDate" Visible="False" />
                <asp:BoundField DataField="ModifyMan" HeaderText="ModifyMan" SortExpression="ModifyMan" Visible="False" />
                <asp:BoundField DataField="ModifyManName" HeaderText="ModifyManName" ReadOnly="True" SortExpression="ModifyManName" Visible="False" />
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
        <asp:FormView ID="fvExceedSpeedReport_Detail" runat="server" DataKeyNames="CaseNo" DataSourceID="sdsExceedSpeedReport_Detail" Width="100%"
            OnDataBound="fvExceedSpeedReport_Detail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Update" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" qid="120px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseDate_Edit" runat="server" CssClass="text-Right-Blue" Text="日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eCaseDateS_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDateS","{0:yyyy/MM/dd}") %>' Width="40%" />
                                    <asp:Label ID="lbSplit_4" runat="server" CssClass="text-Left-Black" Text="～" />
                                    <asp:Label ID="eCaseDateE_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDateE","{0:yyyy/MM/dd}") %>' Width="40%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Edit_TextChanged" Text='<%# Bind("DepNo") %>' Width="25%" />
                                    <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="65%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarID_Edit" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Car_ID") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_Edit" runat="server" Text='<%# Eval("CaseNo") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDRiver_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDriver_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Edit_TextChanged" Text='<%# Bind("Driver") %>' Width="35%" />
                                    <asp:Label ID="eDriverName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DriverName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbException_Edit" runat="server" CssClass="text-Right-Blue" Text="異常項目：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlException_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlException_Edit_SelectedIndexChanged" Width="50px">
                                        <asp:ListItem Value="無" Text="無" Selected="True" />
                                        <asp:ListItem Value="有" Text="有" />
                                    </asp:DropDownList>
                                    <asp:Label ID="eException_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Exception") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAbnormalValue_Edit" runat="server" CssClass="text-Right-Blue" Text="異常數值：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAbnormalValue_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("AbnormalValue") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAttachment_Edit" runat="server" CssClass="text-Right-Blue" Text="附件編號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAttachment_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Attachment") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbInspector_Edit" runat="server" CssClass="text-Right-Blue" Text="查核人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eInspector_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Inspector") %>' Width="35%" />
                                    <asp:Label ID="eInspectorName_Edit" runat="server" CssClass="text-Left-Black" Width="55%" />
                                </td>
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
                                    <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-8Col" colspan="3">
                                    <asp:Label ID="eCaseYM_Edit" runat="server" Text='<%# Eval("CaseYM") %>' Visible="false" />
                                    <asp:Label ID="eMonthStep_Edit" runat="server" Text='<%# Eval("MonthStep") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("ModifyMan") %>' Width="35%" />
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
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Insert" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCaseDate_INS" runat="server" CssClass="text-Right-Blue" Text="日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eCaseDateS_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDateS","{0:yyyy/MM/dd}") %>' Width="40%" />
                                    <asp:Label ID="lbSplit_5" runat="server" CssClass="text-Left-Black" Text="～" />
                                    <asp:Label ID="eCaseDateE_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDateE","{0:yyyy/MM/dd}") %>' Width="40%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Bind("DepNo") %>' Width="25%" />
                                    <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="65%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCarID_INS" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Car_ID") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColWidth-8Col">
                                    <asp:Label ID="eCaseNo_INS" runat="server" Text='<%# Eval("CaseNo") %>' Visible="false" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDRiver_INS" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDriver_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_INS_TextChanged" Text='<%# Bind("Driver") %>' Width="35%" />
                                    <asp:Label ID="eDriverName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DriverName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbException_INS" runat="server" CssClass="text-Right-Blue" Text="異常項目：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlException_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlException_INS_SelectedIndexChanged" Width="50px">
                                        <asp:ListItem Value="無" Text="無" Selected="True" />
                                        <asp:ListItem Value="有" Text="有" />
                                    </asp:DropDownList>
                                    <asp:Label ID="eException_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Exception") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAbnormalValue_INS" runat="server" CssClass="text-Right-Blue" Text="異常數值：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAbnormalValue_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("AbnormalValue") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAttachment_INS" runat="server" CssClass="text-Right-Blue" Text="附件編號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eAttachment_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Attachment") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Height="95%" Width="98%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbInspector_INS" runat="server" CssClass="text-Right-Blue" Text="查核人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eInspector_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Inspector") %>' Width="35%" />
                                    <asp:Label ID="eInspectorName_INS" runat="server" CssClass="text-Left-Black" Width="55%" />
                                </td>
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
                                    <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Bind("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-8Col" colspan="3">
                                    <asp:Label ID="eCaseYM_INS" runat="server" Text='<%# Eval("CaseYM") %>' Visible="false" />
                                    <asp:Label ID="eMonthStep_INS" runat="server" Text='<%# Eval("MonthStep") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人：" Width="90%" />
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
            <ItemTemplate>
                <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" Visible="false" />
                &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="Edit" Text="修改" Width="120px" />
                &nbsp;<asp:Button ID="bbDelete_List" runat="server" CausesValidation="False" CssClass="button-Red" OnClick="bbDelete_List_Click" Text="刪除" Width="120px" Visible="false" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCaseDate_List" runat="server" CssClass="text-Right-Blue" Text="日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eCaseDateS_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDateS","{0:yyyy/MM/dd}") %>' Width="40%" />
                            <asp:Label ID="lbSplit_6" runat="server" CssClass="text-Left-Black" Text="～" />
                            <asp:Label ID="eCaseDateE_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CaseDateE","{0:yyyy/MM/dd}") %>' Width="40%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCarID_List" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCar_ID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColWidth-8Col">
                            <asp:Label ID="eCaseNo" runat="server" Text='<%# Eval("CaseNo") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDRiver_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDriver_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver") %>' Width="35%" />
                            <asp:Label ID="eDriverName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbException_List" runat="server" CssClass="text-Right-Blue" Text="異常項目：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eException_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Exception") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAbnormalValue_List" runat="server" CssClass="text-Right-Blue" Text="異常數值：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAbnormalValue_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AbnormalValue") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAttachment_List" runat="server" CssClass="text-Right-Blue" Text="附件編號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eAttachment_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Attachment") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Height="95%" Width="98%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbInspector_List" runat="server" CssClass="text-Right-Blue" Text="查核人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eInspector_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Inspector") %>' Width="35%" />
                            <asp:Label ID="eInspectorName_List" runat="server" CssClass="text-Left-Black" Width="55%" />
                        </td>
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
                            <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-8Col" colspan="3">
                            <asp:Label ID="eCaseYM_List" runat="server" Text='<%# Eval("CaseYM") %>' Visible="false" />
                            <asp:Label ID="eMonthStep_List" runat="server" Text='<%# Eval("MonthStep") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="90%" />
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
            <LocalReport ReportPath="Report\ExceedSpeedReportP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsExceedSpeedReport_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, CaseYM, MonthStep, CaseDateS, CaseDateE, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DepNo)) AS DepName, Car_ID, Driver, DriverName, Exception, AbnormalValue, Attachment, Remark, Inspector, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = e.Inspector)) AS InspectorName, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = e.BuMan)) AS BuManName, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = e.ModifyMan)) AS ModifyManName FROM ExceedSpeedReport AS e WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsExceedSpeedReport_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        DeleteCommand="DELETE FROM ExceedSpeedReport WHERE (CaseNo = @CaseNo)"
        InsertCommand="INSERT INTO ExceedSpeedReport(CaseNo, CaseYM, MonthStep, CaseDateS, CaseDateE, DepNo, Car_ID, Driver, DriverName, Exception, AbnormalValue, Attachment, Remark, Inspector, BuDate, BuMan) VALUES (@CaseNo, @CaseYM, @MonthStep, @CaseDateS, @CaseDateE, @DepNo, @Car_ID, @Driver, @DriverName, @Exception, @AbnormalValue, @Attachment, @Remark, @Inspector, @BuDate, @BuMan)"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CaseNo, CaseYM, MonthStep, CaseDateS, CaseDateE, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DepNo)) AS DepName, Car_ID, Driver, DriverName, Exception, AbnormalValue, Attachment, Remark, Inspector, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = e.Inspector)) AS InspectorName, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = e.BuMan)) AS BuManName, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = e.ModifyMan)) AS ModifyManName FROM ExceedSpeedReport AS e WHERE (CaseNo = @CaseNo)"
        UpdateCommand="UPDATE ExceedSpeedReport SET DepNo = @DepNo, Driver = @Driver, DriverName = @DriverName, Exception = @Exception, AbnormalValue = @AbnormalValue, Attachment = @Attachment, Remark = @Remark, Inspector = @Inspector, ModifyDate = @ModifyDate, ModifyMan = @ModifyMan WHERE (CaseNo = @CaseNo)"
        OnDeleted="sdsExceedSpeedReport_Detail_Deleted"
        OnInserted="sdsExceedSpeedReport_Detail_Inserted"
        OnInserting="sdsExceedSpeedReport_Detail_Inserting"
        OnUpdated="sdsExceedSpeedReport_Detail_Updated"
        OnUpdating="sdsExceedSpeedReport_Detail_Updating">
        <DeleteParameters>
            <asp:Parameter Name="CaseNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="CaseNo" />
            <asp:Parameter Name="CaseYM" />
            <asp:Parameter Name="MonthStep" />
            <asp:Parameter Name="CaseDateS" />
            <asp:Parameter Name="CaseDateE" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="DriverName" />
            <asp:Parameter Name="Exception" />
            <asp:Parameter Name="AbnormalValue" />
            <asp:Parameter Name="Attachment" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="Inspector" />
            <asp:Parameter Name="BuDate" />
            <asp:Parameter Name="BuMan" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridExceedSpeedReport_List" Name="CaseNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="DriverName" />
            <asp:Parameter Name="Exception" />
            <asp:Parameter Name="AbnormalValue" />
            <asp:Parameter Name="Attachment" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="Inspector" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="CaseNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
</asp:Content>
