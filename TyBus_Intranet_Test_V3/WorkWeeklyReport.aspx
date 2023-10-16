<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="WorkWeeklyReport.aspx.cs" Inherits="TyBus_Intranet_Test_V3.WorkWeeklyReport" %>

<asp:Content ID="WorkWeeklyReportForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">工作週報上傳作業</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="單位：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_Search_TextChanged" Width="30%" />
                    <asp:Label ID="eDepName_Search" runat="server" CssClass="text-Left-Black" Width="65%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbEmpNo_Search" runat="server" CssClass="text-Right-Blue" Text="上傳人員：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eEmpNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eEmpNo_Search_TextChanged" Width="30%" />
                    <asp:Label ID="eEmpName_Search" runat="server" CssClass="text-Left-Black" Width="65%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbBuDate_Search" runat="server" CssClass="text-Right-Blue" Text="建立日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eBuDate_S_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eBuDate_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbUploadDate_Search" runat="server" CssClass="text-Right-Blue" Text="上傳日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eUploadDate_S_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="～" />
                    <asp:TextBox ID="eUploadDate_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col" colspan="2" />
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="搜尋" OnClick="bbSearch_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbExcel_Search" runat="server" CssClass="button-Black" OnClick="bbExcel_Search_Click" Text="匯出 EXCEL" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClear_Search" runat="server" CssClass="button-Blue" Text="清除" OnClick="bbClear_Search_Click" Width="95%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="離開" OnClick="bbClose_Click" Width="95%" />
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
        <asp:GridView ID="gridWorkWeeklyReport" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="#999999" BorderStyle="None" BorderWidth="1px" CellPadding="3" DataKeyNames="IndexNo" DataSourceID="sdsWorkWeeklyReport_List" GridLines="Vertical" PageSize="5" Width="100%">
            <AlternatingRowStyle BackColor="#DCDCDC" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="IndexNo" HeaderText="序號" ReadOnly="True" SortExpression="IndexNo" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="單位" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="EmpNo" HeaderText="EmpNo" SortExpression="EmpNo" Visible="False" />
                <asp:BoundField DataField="EmpName" HeaderText="姓名" ReadOnly="True" SortExpression="EmpName" />
                <asp:BoundField DataField="Title" HeaderText="Title" SortExpression="Title" Visible="False" />
                <asp:BoundField DataField="Title_C" HeaderText="職稱" ReadOnly="True" SortExpression="Title_C" />
                <asp:BoundField DataField="BuDate" DataFormatString="{0:D}" HeaderText="建立日期" SortExpression="BuDate" />
                <asp:BoundField DataField="BuMan" HeaderText="BuMan" SortExpression="BuMan" Visible="False" />
                <asp:BoundField DataField="BuManName" HeaderText="BuManName" ReadOnly="True" SortExpression="BuManName" Visible="False" />
                <asp:BoundField DataField="UploadDate" DataFormatString="{0:D}" HeaderText="上傳日期" SortExpression="UploadDate" />
                <asp:BoundField DataField="UploadTime" HeaderText="上傳時間" SortExpression="UploadTime" />
                <asp:BoundField DataField="FilePath" HeaderText="FilePath" SortExpression="FilePath" Visible="False" />
                <asp:BoundField DataField="Remark" HeaderText="Remark" SortExpression="Remark" Visible="False" />
                <asp:BoundField DataField="ModifyDate" HeaderText="ModifyDate" SortExpression="ModifyDate" Visible="False" />
                <asp:BoundField DataField="ModifyMan" HeaderText="ModifyMan" SortExpression="ModifyMan" Visible="False" />
                <asp:BoundField DataField="ModifyManName" HeaderText="ModifyManName" ReadOnly="True" SortExpression="ModifyManName" Visible="False" />
            </Columns>
            <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
            <HeaderStyle BackColor="#000084" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <RowStyle BackColor="#EEEEEE" ForeColor="Black" />
            <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F1F1F1" />
            <SortedAscendingHeaderStyle BackColor="#0000A9" />
            <SortedDescendingCellStyle BackColor="#CAC9C9" />
            <SortedDescendingHeaderStyle BackColor="#000065" />
        </asp:GridView>
        <asp:FormView ID="fvWorkWeeklyReport" runat="server" DataKeyNames="IndexNo" DataSourceID="sdsWorkWeeklyReport_Detail" Width="100%" OnDataBound="fvWorkWeeklyReport_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_Edit_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="lbIndexNo_Edit" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col" colspan="8">
                                    <asp:Label ID="eIndexNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("IndexNo") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="單位：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                                    <asp:Label ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="30%" />
                                    <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="65%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="lbTitle_Edit" runat="server" CssClass="text-Right-Blue" Text="職稱：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                                    <asp:Label ID="eTitle_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title") %>' Width="30%" />
                                    <asp:Label ID="eTitle_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title_C") %>' Width="65%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="lbEmpNo_Edit" runat="server" CssClass="text-Right-Blue" Text="姓名：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                                    <asp:Label ID="eEmpNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="30%" />
                                    <asp:Label ID="eEmpName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="65%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建立日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                                    <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                                    <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="65%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="lbUploadDate_Edit" runat="server" CssClass="text-Right-Blue" Text="上傳日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="eUploadDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UploadDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="lbUploadTime_Edit" runat="server" CssClass="text-Right-Blue" Text="上傳時間：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="eUploadTime_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UploadTime") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="lbFilePath_Edit" runat="server" CssClass="text-Right-Blue" Text="附件檔案：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col" colspan="8">
                                    <asp:Label ID="eFilePath_Edit" runat="server" CssClass="text-Left-Red" Text='<%# Eval("FilePath") %>' Width="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine_Low ColBorder ColWidth-9Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                                </td>
                                <td class="MultiLine_Low ColBorder ColWidth-9Col" colspan="8">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="99%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-9Col" colspan="4"></td>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="lbModdifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="30%" />
                                    <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="65%" />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_INS_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbIndexNo_INS" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="8">
                            <asp:Label ID="eIndexNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("IndexNo") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="單位：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                            <asp:Label ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="30%" />
                            <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="65%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbTitle_INS" runat="server" CssClass="text-Right-Blue" Text="職稱：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                            <asp:Label ID="eTitle_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title") %>' Width="30%" />
                            <asp:Label ID="eTitle_C_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title_C") %>' Width="65%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbEmpNo_INS" runat="server" CssClass="text-Right-Blue" Text="姓名：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                            <asp:Label ID="eEmpNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="30%" />
                            <asp:Label ID="eEmpName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="65%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建立日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                            <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                            <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="65%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbUploadDate_INS" runat="server" CssClass="text-Right-Blue" Text="上傳日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="eUploadDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UploadDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbUploadTime_INS" runat="server" CssClass="text-Right-Blue" Text="上傳時間：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="eUploadTime_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UploadTime") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbFilePath_INS" runat="server" CssClass="text-Right-Blue" Text="附件檔案：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="8">
                            <asp:FileUpload ID="fuFilePath_INS" runat="server" CssClass="text-Left-Blue" Width="95%" />
                            <asp:Label ID="eFilePath_INS" runat="server" Text='<%# Eval("FilePath") %>' Visible="false" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_Low ColBorder ColWidth-9Col">
                            <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                        </td>
                        <td class="MultiLine_Low ColBorder ColWidth-9Col" colspan="8">
                            <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="99%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-9Col" colspan="4"></td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbModdifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                            <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="30%" />
                            <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="65%" />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </InsertItemTemplate>
            <ItemTemplate>
                <asp:Button ID="bbNew" runat="server" CssClass="button-Black" CausesValidation="False" CommandName="New" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit" runat="server" CssClass="button-Blue" CausesValidation="False" CommandName="Edit" Text="編輯" Width="120px" />
                &nbsp;<asp:Button ID="bbDelete" runat="server" CssClass="button-Red" OnClick="bbDelete_Click" Text="刪除" Width="120px" />
                &nbsp;<asp:Button ID="bbDownload" runat="server" CssClass="button-Blue" CausesValidation="false" OnClick="bbDownload_Click" Text="下載附檔" Width="120px" />
                &nbsp;<asp:Button ID="bbReupload" runat="server" CssClass="button-Red" CausesValidation="false" OnClick="bbReupload_Click" Text="重新上傳" Width="120px" />
                &nbsp;<asp:FileUpload ID="fuReupload" runat="server" CssClass="text-Left-Blue" Width="200px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbIndexNo_List" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="8">
                            <asp:Label ID="eIndexNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("IndexNo") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="單位：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="30%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="65%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbTitle_List" runat="server" CssClass="text-Right-Blue" Text="職稱：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                            <asp:Label ID="eTitle_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title") %>' Width="30%" />
                            <asp:Label ID="eTitle_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Title_C") %>' Width="65%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbEmpNo_List" runat="server" CssClass="text-Right-Blue" Text="姓名：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                            <asp:Label ID="eEmpNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpNo") %>' Width="30%" />
                            <asp:Label ID="eEmpName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("EmpName") %>' Width="65%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建立日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                            <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                            <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="65%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbUploadDate_List" runat="server" CssClass="text-Right-Blue" Text="上傳日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="eUploadDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UploadDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbUploadTime_List" runat="server" CssClass="text-Right-Blue" Text="上傳時間：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="eUploadTime_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UploadTime") %>' Width="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbFilePath_List" runat="server" CssClass="text-Right-Blue" Text="附件檔案：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="8">
                            <asp:Label ID="eFilePath_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FilePath") %>' Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine_Low ColBorder ColWidth-9Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="95%" />
                        </td>
                        <td class="MultiLine_Low ColBorder ColWidth-9Col" colspan="8">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Width="99%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-9Col" colspan="4"></td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbModdifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate", "{0:yyyy/MM/dd}") %>' Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                            <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="30%" />
                            <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="65%" />
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
            </ItemTemplate>
            <EmptyDataTemplate>
                <asp:Button ID="bbNew_Empty" runat="server" CssClass="button-Black" CausesValidation="false" CommandName="New" Text="新增" Width="120px" />
            </EmptyDataTemplate>
        </asp:FormView>
        <asp:SqlDataSource ID="sdsWorkWeeklyReport_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
            SelectCommand="SELECT IndexNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, EmpNo, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.EmpNo)) AS EmpName, Title, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '員工資料        EMPLOYEE        TITLE') AND (CLASSNO = a.Title)) AS Title_C, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = a.BuMan)) AS BuManName, UploadDate, UploadTime, FilePath, Remark, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName FROM WorkWeeklyReport AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
        <asp:SqlDataSource ID="sdsWorkWeeklyReport_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
            SelectCommand="SELECT IndexNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, EmpNo, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.EmpNo)) AS EmpName, Title, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '員工資料        EMPLOYEE        TITLE') AND (CLASSNO = a.Title)) AS Title_C, BuDate, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = a.BuMan)) AS BuManName, UploadDate, UploadTime, FilePath, Remark, ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName FROM WorkWeeklyReport AS a WHERE (IndexNo = @IndexNo)" DeleteCommand="DELETE FROM WorkWeeklyReport WHERE (IndexNo = @IndexNo)" InsertCommand="INSERT INTO WorkWeeklyReport(IndexNo, DepNo, EmpNo, Title, BuDate, BuMan, UploadDate, UploadTime, FilePath, Remark) VALUES (@IndexNo, @DepNo, @EmpNo, @Title, @BuDate, @BuMan, @UploadDate, @UploadTime, @FilePath, @Remark)" UpdateCommand="UPDATE WorkWeeklyReport SET Remark = @Remark, ModifyDate = @ModifyDate, ModifyMan = @ModifyMan WHERE (IndexNo = @IndexNo)">
            <DeleteParameters>
                <asp:Parameter Name="IndexNo" />
            </DeleteParameters>
            <InsertParameters>
                <asp:Parameter Name="IndexNo" />
                <asp:Parameter Name="DepNo" />
                <asp:Parameter Name="EmpNo" />
                <asp:Parameter Name="Title" />
                <asp:Parameter Name="BuDate" />
                <asp:Parameter Name="BuMan" />
                <asp:Parameter Name="UploadDate" />
                <asp:Parameter Name="UploadTime" />
                <asp:Parameter Name="FilePath" />
                <asp:Parameter Name="Remark" />
            </InsertParameters>
            <SelectParameters>
                <asp:ControlParameter ControlID="gridWorkWeeklyReport" Name="IndexNo" PropertyName="SelectedValue" />
            </SelectParameters>
            <UpdateParameters>
                <asp:Parameter Name="Remark" />
                <asp:Parameter Name="ModifyDate" />
                <asp:Parameter Name="ModifyMan" />
                <asp:Parameter Name="IndexNo" />
            </UpdateParameters>
        </asp:SqlDataSource>
    </asp:Panel>
</asp:Content>
