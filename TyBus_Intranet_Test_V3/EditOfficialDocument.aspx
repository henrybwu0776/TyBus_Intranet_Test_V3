<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EditOfficialDocument.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EditOfficialDocument" %>

<asp:Content ID="EditOfficialDocumentForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">公文登錄資料修改</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" GroupingText="查詢條件" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDocDep_Search" runat="server" CssClass="text-Right-Blue" Text="收發單位：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDocDep_Start_Search" runat="server" CssClass="text-Left-Black" OnTextChanged="eDocDep_Start_Search_TextChanged" AutoPostBack="true" Width="40%" />
                    <asp:Label ID="lbSplit1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDocDep_End_Search" runat="server" CssClass="text-Left-Black" OnTextChanged="eDocDep_End_Search_TextChanged" AutoPostBack="true" Width="40%" />
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
                        DataSourceID="sdsDocType_Search" DataTextField="ClassTxt" DataValueField="ClassNo" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlDocType_Search_SelectedIndexChanged">
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
                    <asp:TextBox ID="eUndertaker_Start_Search" runat="server" CssClass="text-Left-Black" OnTextChanged="eUndertaker_Start_Search_TextChanged" AutoPostBack="true" Width="40%" />
                    <asp:Label ID="lbSplit4" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eUndertaker_End_Search" runat="server" CssClass="text-Left-Black" OnTextChanged="eUndertaker_End_Search_TextChanged" AutoPostBack="true" Width="40%" />
                    <br />
                    <asp:Label ID="eUndertakerName_Start_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit5" runat="server" Width="5%" />
                    <asp:Label ID="eUndertakerName_End_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDocTitle_Search" runat="server" CssClass="text-Right-Blue" Text="事由：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDocTitle_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDocNo_Search" runat="server" CssClass="text-Right-Blue" Text="文號：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eDocNo_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
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
    <asp:Panel ID="plDataList" runat="server" GroupingText="公文列表" CssClass="ShowPanel">
        <asp:FormView ID="fvEditOfficialDocument" runat="server" Width="100%" DataKeyNames="DocIndex" DataSourceID="sdsOfficialDocument_Detail"
            OnDataBound="fvEditOfficialDocument_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CssClass="button-Black" CausesValidation="True" CommandName="Update" Text="更新" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDocNo_Edit" runat="server" CssClass="text-Right-Blue" Text="收發文號：" Width="90%" />
                                    <asp:Label ID="eDocIndex_Edit" runat="server" Text='<%# Eval("DocIndex") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eDocYears_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocYears") %>' />
                                    <asp:Label ID="lbSplit_Edit1" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                                    <asp:Label ID="eDocFirstWord_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocFirstWord_C") %>' />
                                    <asp:Label ID="lbSplit_Edit2" runat="server" CssClass="text-Left-Black" Text=" 字第 " />
                                    <asp:Label ID="eDocNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocNo") %>' />
                                    <asp:Label ID="lbSplit_Edit3" runat="server" CssClass="text-Left-Black" Text=" 號" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDocDate_Edit" runat="server" CssClass="text-Right-Blue" Text="收發文日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDocDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDocDep_Edit" runat="server" CssClass="text-Right-Blue" Text="收發文單位：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eDocDep_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocDep_C") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDocType_Edit" runat="server" CssClass="text-Right-Blue" Text="文別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eDocType_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DocType") %>' />
                                    <asp:Label ID="eDocType_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocType_C") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbUndertaker_Edit" runat="server" CssClass="text-Right-Blue" Text="承辦人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eUndertaker_Edit" runat="server" CssClass="text-Left-Black" OnTextChanged="eUndertaker_Edit_TextChanged" Text='<%# Bind("Undertaker") %>' Width="30%" />
                                    <asp:Label ID="eUndertaker_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Undertaker_C") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbAttachement_Edit" runat="server" CssClass="text-Right-Blue" Text="收發文附件：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eAttachement_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Attachement") %>' Height="95%" Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDocSourceUnit_Edit" runat="server" CssClass="text-Right-Blue" Text="來文機關：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eDocSourceUnit_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("DocSourceUnit") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbOutsideDocNo_Edit" runat="server" CssClass="text-Right-Blue" Text="來文字號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eOutsideDocFirstWord_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("OutsideDocFirstWord") %>' Width="20%" />
                                    <asp:Label ID="lbSplit_Edit4" runat="server" CssClass="text-Left-Black" Text=" 字第 " />
                                    <asp:TextBox ID="eOutsideDocNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("OutsideDocNo") %>' Width="30%" />
                                    <asp:Label ID="lbSplit_Edit5" runat="server" CssClass="text-Left-Black" Text=" 號" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDocTitle_Edit" runat="server" CssClass="text-Right-Blue" Text="事由：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eDocTitle_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("DocTitle") %>' Height="95%" Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbImplementation_Edit" runat="server" CssClass="text-Right-Blue" Text="辦理情形：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eImplementation_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Implementation") %>' Height="95%" Width="97%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備考：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Bind("Remark") %>' Height="95%" Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbStoreDate_Edit" runat="server" CssClass="text-Right-Blue" Text="歸檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eStoreDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StoreDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbStoreMan_Edit" runat="server" CssClass="text-Right-Blue" Text="歸檔人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eStoreMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StoreMan") %>' Width="90%" />
                                    <br />
                                    <asp:Label ID="eStoreMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StoreMan_C") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_Store_Edit" runat="server" CssClass="text-Right-Blue" Text="歸檔備考：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="3">
                                    <asp:TextBox ID="eRemark_Store_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Bind("Remark_Store") %>' Height="95%" Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuildMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan_C") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuildDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuildDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("BuildDate","{0:yyyy/MM/dd}") %>' Width="90%" />
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
            <ItemTemplate>
                <asp:Button ID="bbEdit_List" runat="server" CssClass="button-Black" CausesValidation="False" CommandName="Edit" Text="編輯" />
                &nbsp;<asp:Button ID="bbDelete_List" runat="server" CssClass="button-Red" CausesValidation="False" OnClick="bbDelete_List_Click" Text="刪除" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDocNo_List" runat="server" CssClass="text-Right-Blue" Text="收發文號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDocYears_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocYears") %>' />
                            <asp:Label ID="lbSplit_List1" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                            <asp:Label ID="eDocFirstWord_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocFirstWord_C") %>' />
                            <asp:Label ID="lbSplit_List2" runat="server" CssClass="text-Left-Black" Text=" 字第 " />
                            <asp:Label ID="eDocNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocNo") %>' />
                            <asp:Label ID="lbSplit_List3" runat="server" CssClass="text-Left-Black" Text=" 號" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDocDate_List" runat="server" CssClass="text-Right-Blue" Text="收發文日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDocDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocDate","{0:yyyy/MM/dd}") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDocDep_List" runat="server" CssClass="text-Right-Blue" Text="收發文單位：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDocDep_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocDep") %>' Width="30%" />
                            <asp:Label ID="eDocDep_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocDep_C") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDocType_List" runat="server" CssClass="text-Right-Blue" Text="文別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDocType_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocType") %>' Width="45%" />
                            <asp:Label ID="eDocType_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocType_C") %>' Width="45%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbUndertaker_List" runat="server" CssClass="text-Right-Blue" Text="承辦人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eUndertaker_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Undertaker") %>' Width="30%" />
                            <asp:Label ID="eUndertaker_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Undertaker_C") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbAttachement_List" runat="server" CssClass="text-Right-Blue" Text="收發文附件：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="2">
                            <asp:TextBox ID="eAttachement_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Attachement") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDocSourceUnit_List" runat="server" CssClass="text-Right-Blue" Text="來文機關：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eDocSourceUnit_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DocSourceUnit") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbOutsideDocNo_List" runat="server" CssClass="text-Right-Blue" Text="來文字號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                            <asp:Label ID="eOutsideDocFirstWord_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OutsideDocFirstWord") %>' Width="20%" />
                            <asp:Label ID="lbSplit_List4" runat="server" CssClass="text-Left-Black" Text=" 字第 " />
                            <asp:Label ID="eOutsideDocNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("OutsideDocNo") %>' Width="30%" />
                            <asp:Label ID="lbSplit_List5" runat="server" CssClass="text-Left-Black" Text=" 號" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDocTitle_List" runat="server" CssClass="text-Right-Blue" Text="事由：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eDocTitle_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("DocTitle") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbImplementation_List" runat="server" CssClass="text-Right-Blue" Text="辦理情形：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="3">
                            <asp:TextBox ID="eImplementation_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Implementation") %>' Height="95%" Width="97%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備考：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="3">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStoreDate_List" runat="server" CssClass="text-Right-Blue" Text="歸檔日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStoreDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StoreDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbStoreMan_List" runat="server" CssClass="text-Right-Blue" Text="歸檔人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eStoreMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("StoreMan_C") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_Store_List" runat="server" CssClass="text-Right-Blue" Text="歸檔備考：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="3">
                            <asp:TextBox ID="eRemark_Store_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark_Store") %>' Height="95%" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuildMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan") %>' Width="30%" />
                            <asp:Label ID="eBuildMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildMan_C") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuildDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuildDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuildDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColWidth-8Col">
                            <asp:Label ID="eDocIndex_List" runat="server" Text='<%# Eval("DocIndex") %>' Visible="false" />
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
        <asp:GridView ID="gridDataList" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White"
            BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1"
            DataKeyNames="DocIndex" DataSourceID="sdsOfficialDocument_List" GridLines="None" Width="100%" PageSize="5">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="DocIndex" HeaderText="DocIndex" ReadOnly="True" SortExpression="DocIndex" Visible="False" />
                <asp:TemplateField HeaderText="文號">
                    <ItemTemplate>
                        <asp:Label ID="DocYears" runat="server" Text='<%# Eval("DocYears") %>' />
                        <asp:Label ID="DocNoSplit1" runat="server" Text=" 年 " />
                        <asp:Label ID="DocFirstCWord" runat="server" Text='<%# Eval("DocFirstWord_C") %>' />
                        <asp:Label ID="DocNoSplit2" runat="server" Text="  字第  " />
                        <asp:Label ID="DocNo" runat="server" Text='<%# Eval("DocNo") %>' />
                        <asp:Label ID="DocNoSplit3" runat="server" Text=" 號 " />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="DocDate" DataFormatString="{0:d}" HeaderText="收發文日期" SortExpression="DocDate" />
                <asp:BoundField DataField="DocDep" HeaderText="DocDep" SortExpression="DocDep" Visible="False" />
                <asp:BoundField DataField="DocDep_C" HeaderText="收發文單位" ReadOnly="True" SortExpression="DocDep_C" />
                <asp:BoundField DataField="DocFirstWord" HeaderText="DocFirstWord" SortExpression="DocFirstWord" Visible="False" />
                <asp:BoundField DataField="DocNo" HeaderText="DocNo" SortExpression="DocNo" Visible="False" />
                <asp:BoundField DataField="DocSourceUnit" HeaderText="來文機關" SortExpression="DocSourceUnit" />
                <asp:BoundField DataField="DocType" HeaderText="DocType" SortExpression="DocType" Visible="False" />
                <asp:BoundField DataField="DocType_C" HeaderText="文別" ReadOnly="True" SortExpression="DocType_C" />
                <asp:BoundField DataField="DocTitle" HeaderText="事由" SortExpression="DocTitle" />
                <asp:BoundField DataField="Undertaker" HeaderText="Undertaker" SortExpression="Undertaker" Visible="False" />
                <asp:BoundField DataField="Undertaker_C" HeaderText="承辦人" ReadOnly="True" SortExpression="Undertaker_C" />
                <asp:BoundField DataField="OutsideDocFirstWord" HeaderText="OutsideDocFirstWord" SortExpression="OutsideDocFirstWord" Visible="False" />
                <asp:BoundField DataField="OutsideDocNo" HeaderText="OutsideDocNo" SortExpression="OutsideDocNo" Visible="False" />
                <asp:BoundField DataField="Attachement" HeaderText="附件" SortExpression="Attachement" />
                <asp:BoundField DataField="Implementation" HeaderText="辦理情形" SortExpression="Implementation" />
                <asp:BoundField DataField="BuildMan" HeaderText="BuildMan" SortExpression="BuildMan" Visible="False" />
                <asp:BoundField DataField="BuildMan_C" HeaderText="建檔人" ReadOnly="True" SortExpression="BuildMan_C" />
                <asp:BoundField DataField="BuildDate" DataFormatString="{0:d}" HeaderText="建檔日期" SortExpression="BuildDate" />
                <asp:BoundField DataField="Remark" HeaderText="收發文備註" SortExpression="Remark" />
                <asp:BoundField DataField="StoreDate" DataFormatString="{0:d}" HeaderText="歸檔日期" SortExpression="StoreDate" />
                <asp:BoundField DataField="StoreMan" HeaderText="StoreMan" SortExpression="StoreMan" Visible="False" />
                <asp:BoundField DataField="StoreMan_C" HeaderText="歸檔人" ReadOnly="True" SortExpression="StoreMan_C" />
                <asp:BoundField DataField="Remark_Store" HeaderText="歸檔備考" SortExpression="Remark_Store" />
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
    </asp:Panel>
    <asp:SqlDataSource ID="sdsOfficialDocument_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT DocIndex, DocDate, DocYears, DocDep, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DocDep)) AS DocDep_C, DocFirstWord, (SELECT DocFirstCWord FROM DOCFirstWord WHERE (FWNo = a.DocFirstWord)) AS DocFirstWord_C, DocNo, DocSourceUnit, DocType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = CAST('公文收發登記' AS char(16)) + CAST('OfficialDocument' AS char(16)) + 'DocType') AND (CLASSNO = a.DocType)) AS DocType_C, DocTitle, Undertaker, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.Undertaker)) AS Undertaker_C, OutsideDocFirstWord, OutsideDocNo, Attachement, Implementation, BuildMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = a.BuildMan)) AS BuildMan_C, BuildDate, Remark, StoreDate, StoreMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.StoreMan)) AS StoreMan_C, Remark_Store FROM OfficialDocument AS a"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsDocType_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = CAST('公文收發登記' AS char(16)) + CAST('OfficialDocument' AS char(16)) + 'DocType') ORDER BY ClassNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsDocType_Edit" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = CAST('公文收發登記' AS char(16)) + CAST('OfficialDocument' AS char(16)) + 'DocType') ORDER BY ClassNo"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsOfficialDocument_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT DocIndex, DocDate, DocYears, DocDep, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DocDep)) AS DocDep_C, DocFirstWord, (SELECT DocFirstCWord FROM DOCFirstWord WHERE (FWNo = a.DocFirstWord)) AS DocFirstWord_C, DocNo, DocSourceUnit, DocType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = CAST('公文收發登記' AS char(16)) + CAST('OfficialDocument' AS char(16)) + 'DocType') AND (CLASSNO = a.DocType)) AS DocType_C, DocTitle, Undertaker, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.Undertaker)) AS Undertaker_C, OutsideDocFirstWord, OutsideDocNo, Attachement, Implementation, BuildMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = a.BuildMan)) AS BuildMan_C, BuildDate, Remark, StoreDate, StoreMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.StoreMan)) AS StoreMan_C, Remark_Store FROM OfficialDocument AS a WHERE (DocIndex = @DocIndex)"
        DeleteCommand="DELETE FROM OfficialDocument WHERE (DocIndex = @DocIndex)"
        UpdateCommand="UPDATE OfficialDocument SET DocDate = @DocDate, DocType = @DocType, DocTitle = @DocTitle, Undertaker = @Undertaker, Attachement = @Attachement, Implementation = @Implementation, Remark = @Remark, StoreDate = @StoreDate, Remark_Store = @Remark_Store, DocSourceUnit = @DocSourceUnit, OutsideDocFirstWord = @OutsideDocFirstWord, OutsideDocNo = @OutsideDocNo WHERE (DocIndex = @DocIndex)"
        OnDeleted="sdsOfficialDocument_Detail_Deleted"
        OnUpdated="sdsOfficialDocument_Detail_Updated"
        OnUpdating="sdsOfficialDocument_Detail_Updating">
        <DeleteParameters>
            <asp:Parameter Name="DocIndex" />
        </DeleteParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridDataList" Name="DocIndex" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="DocDate" />
            <asp:Parameter Name="DocType" />
            <asp:Parameter Name="DocTitle" />
            <asp:Parameter Name="Undertaker" />
            <asp:Parameter Name="Attachement" />
            <asp:Parameter Name="Implementation" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="StoreDate" />
            <asp:Parameter Name="Remark_Store" />
            <asp:Parameter Name="DocIndex" />
            <asp:Parameter Name="DocSourceUnit" />
            <asp:Parameter Name="OutsideDocFirstWord" />
            <asp:Parameter Name="OutsideDocNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
</asp:Content>
