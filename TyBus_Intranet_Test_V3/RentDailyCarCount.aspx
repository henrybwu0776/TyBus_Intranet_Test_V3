<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="RentDailyCarCount.aspx.cs" Inherits="TyBus_Intranet_Test_V3.RentDailyCarCount" %>

<asp:Content ID="RentDailyCarCountForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">每日車輛可用數</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbCalYM_Search" runat="server" CssClass="text-Right-Blue" Text="計算年月：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eCalYear_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:TextBox ID="eCalMonth_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text=" 月" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:DropDownList ID="ddlDepNo_Search" runat="server" CssClass="text-Left-Black" Width="90%"
                        DataSourceID="sdsDepNo_Search" DataTextField="DepName" DataValueField="DEPNO"
                        AutoPostBack="true" OnSelectedIndexChanged="ddlDepNo_Search_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:Label ID="eDepNo_Search" runat="server" CssClass="text-Left-Black" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbCarCount_Search" runat="server" CssClass="text-Right-Blue" Text="車輛可用數：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:TextBox ID="eCarCount_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbBatchInsert" runat="server" CssClass="button-Blue" Text="批次填寫" OnClick="bbBatchInsert_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
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
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <div class="ShowPanel-Detail_C">
            <asp:GridView ID="gridRentDailyCarCountList" runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="4"
                DataKeyNames="IndexNo" DataSourceID="sdsRentDailyCarCount_List" ForeColor="#333333" GridLines="None" PageSize="5">
                <AlternatingRowStyle BackColor="White" />
                <Columns>
                    <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                    <asp:BoundField DataField="CalDate" DataFormatString="{0:D}" HeaderText="日期" SortExpression="CalDate" />
                    <asp:BoundField DataField="IndexNo" HeaderText="IndexNo" ReadOnly="True" SortExpression="IndexNo" Visible="False" />
                    <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                    <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="True" SortExpression="DepName" />
                    <asp:BoundField DataField="UseableCount" HeaderText="可用車總數" SortExpression="UseableCount" />
                    <asp:BoundField DataField="UseableCount_Red" HeaderText="紅牌車可用數" SortExpression="UseableCount_Red" Visible="False" />
                    <asp:BoundField DataField="UseableCount_Green" HeaderText="綠牌車可用數" SortExpression="UseableCount_Green" Visible="False" />
                    <asp:BoundField DataField="NeedCarCount" HeaderText="當日需求數" SortExpression="NeedCarCount" Visible="False" />
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
        </div>
        <div class="ShowPanel-Detail_B">
            <asp:FormView ID="fvDataDetail" runat="server" DataKeyNames="IndexNo" DataSourceID="sdsRentDailyCarCount_Detail" Width="100%" OnDataBound="fvDataDetail_DataBound">
                <EditItemTemplate>
                    <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_Edit_Click" Text="更新" Width="120px" />
                    &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                    <asp:UpdatePanel ID="upDataEdit" runat="server">
                        <ContentTemplate>
                            <table class="TableSetting">
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbIndexNo_Edit" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="eIndexNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("IndexNo") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbCalDate_Edit" runat="server" CssClass="text-Right-Blue" Text="日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="eCalDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CalDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col" colspan="3">
                                        <asp:Label ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="30%" />
                                        <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbUseableCarCount_Edit" runat="server" CssClass="text-Right-Blue" Text="車輛可用數：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col" colspan="3">
                                        <asp:TextBox ID="eUseableCount_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UseableCount") %>' Width="90%" />
                                        <asp:Label ID="eUseableCountRed_Edit" runat="server" Text='<%# Eval("UseableCount_Red") %>' Visible="false" />
                                        <asp:Label ID="eUseableCountGreen_Edit" runat="server" Text='<%# Eval("UseableCount_Green") %>' Visible="false" />
                                        <asp:Label ID="NeedCarCount_Edit" runat="server" Text='<%# Eval("NeedCarCount") %>' Visible="false" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="MultiLine_Low ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                    </td>
                                    <td class="MultiLine_Low ColBorder ColWidth-4Col" colspan="3">
                                        <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="99%" Height="97%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                                        <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="55%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="30%" />
                                        <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="55%" />
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
                <EmptyDataTemplate>
                    <asp:Button ID="bbNew_Empty" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
                </EmptyDataTemplate>
                <InsertItemTemplate>
                    <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Blue" OnClick="bbOK_INS_Click" Text="確定" Width="120px" />
                    &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                    <asp:UpdatePanel ID="upDataInsert" runat="server">
                        <ContentTemplate>
                            <table class="TableSetting">
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbIndexNo_INS" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="eIndexNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("IndexNo") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbCalDate_INS" runat="server" CssClass="text-Right-Blue" Text="日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="eCalDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CalDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col" colspan="3">
                                        <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Eval("DepNo") %>' Width="30%" />
                                        <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbUseableCarCount_INS" runat="server" CssClass="text-Right-Blue" Text="車輛可用數：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col" colspan="3">
                                        <asp:TextBox ID="eUseableCount_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UseableCount") %>' Width="90%" />
                                        <asp:Label ID="eUseableCountRed_INS" runat="server" Text='<%# Eval("UseableCount_Red") %>' Visible="false" />
                                        <asp:Label ID="eUseableCountGreen_INS" runat="server" Text='<%# Eval("UseableCount_Green") %>' Visible="false" />
                                        <asp:Label ID="NeedCarCount_INS" runat="server" Text='<%# Eval("NeedCarCount") %>' Visible="false" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="MultiLine_Low ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                    </td>
                                    <td class="MultiLine_Low ColBorder ColWidth-4Col" colspan="3">
                                        <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="99%" Height="97%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                                        <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="55%" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                                    </td>
                                    <td class="ColHeight ColBorder ColWidth-4Col">
                                        <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="30%" />
                                        <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="55%" />
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
                <ItemTemplate>
                    <asp:Button ID="bbNew_List" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
                    &nbsp;<asp:Button ID="bbEdit_List" runat="server" CausesValidation="False" CssClass="button-Blue" CommandName="Edit" Text="修改" Width="120px" />
                    &nbsp;<asp:Button ID="bbDelete_List" runat="server" CausesValidation="False" CssClass="button-Red" OnClick="bbDelete_List_Click" Text="刪除" Width="120px" />
                    <table class="TableSetting">
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="lbIndexNo_List" runat="server" CssClass="text-Right-Blue" Text="序號：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="eIndexNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("IndexNo") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="lbCalDate_List" runat="server" CssClass="text-Right-Blue" Text="日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="eCalDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("CalDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-4Col" colspan="3">
                                <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="30%" />
                                <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="lbUseableCarCount_List" runat="server" CssClass="text-Right-Blue" Text="車輛可用數：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-4Col" colspan="3">
                                <asp:Label ID="eUseableCount_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("UseableCount") %>' Width="90%" />
                                <asp:Label ID="eUseableCountRed_List" runat="server" Text='<%# Eval("UseableCount_Red") %>' Visible="false" />
                                <asp:Label ID="eUseableCountGreen_List" runat="server" Text='<%# Eval("UseableCount_Green") %>' Visible="false" />
                                <asp:Label ID="NeedCarCount_List" runat="server" Text='<%# Eval("NeedCarCount") %>' Visible="false" />
                            </td>
                        </tr>
                        <tr>
                            <td class="MultiLine_Low ColBorder ColWidth-4Col">
                                <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                            </td>
                            <td class="MultiLine_Low ColBorder ColWidth-4Col" colspan="3">
                                <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Width="99%" Height="97%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="30%" />
                                <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="55%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人：" Width="90%" />
                            </td>
                            <td class="ColHeight ColBorder ColWidth-4Col">
                                <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="30%" />
                                <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan_C") %>' Width="55%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="ColWidth-4Col" />
                            <td class="ColWidth-4Col" />
                            <td class="ColWidth-4Col" />
                            <td class="ColWidth-4Col" />
                        </tr>
                    </table>
                </ItemTemplate>
            </asp:FormView>
        </div>
    </asp:Panel>
    <asp:Panel ID="plCalender" runat="server" CssClass="ShowPanel-Detail">
        <asp:Calendar ID="calRentDailyCarCount" runat="server" Width="100%" OnDayRender="calRentDailyCarCount_DayRender" OnSelectionChanged="calRentDailyCarCount_SelectionChanged"></asp:Calendar>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsRentDailyCarCount_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT IndexNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, CalDate, UseableCount, UseableCount_Red, UseableCount_Green, NeedCarCount FROM RentDailyCarCount AS a WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsRentDailyCarCount_Detail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" DeleteCommand="DELETE FROM RentDailyCarCount WHERE (IndexNo = @IndexNo)" InsertCommand="INSERT INTO RentDailyCarCount(IndexNo, DepNo, CalDate, UseableCount, Remark, BuMan, BuDate) VALUES (@IndexNo, @DepNo, @CalDate, @UseableCount, @Remark, @BuMan, @BuDate)" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT IndexNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, CalDate, UseableCount, UseableCount_Red, UseableCount_Green, NeedCarCount, Remark, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuMan_C, BuDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyMan_C, ModifyDate FROM RentDailyCarCount AS a WHERE (IndexNo = @IndexNo)" UpdateCommand="UPDATE RentDailyCarCount SET UseableCount = @UseableCount, Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate WHERE (IndexNo = @IndexNo)">
        <DeleteParameters>
            <asp:Parameter Name="IndexNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="IndexNo" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="CalDate" />
            <asp:Parameter Name="UseableCount" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="BuDate" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridRentDailyCarCountList" Name="IndexNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="UseableCount" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="ModifyDate" />
            <asp:Parameter Name="IndexNo" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsDepNo_Search" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT DEPNO, DEPNO + '-' + NAME AS DepName FROM DEPARTMENT WHERE (ISNULL(InSHReport, 'X') = 'V')"></asp:SqlDataSource>
</asp:Content>
