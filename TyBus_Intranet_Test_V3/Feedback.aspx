<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="Feedback.aspx.cs" Inherits="TyBus_Intranet_Test_V3.Feedback" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="FeedbackForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">電腦課派工回報作業</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbBuildMan" runat="server" CssClass="text-Right-Blue" Text="報修申請人：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eBuildMan_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eBuildMan_Search_TextChanged" Width="35%" />
                    <asp:Label ID="eBuildManName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDay_Search" runat="server" CssClass="text-Right-Blue" Text="申請日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eDay_B_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="eSplit_1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eDay_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:RadioButtonList ID="eDisposal_Search" runat="server" RepeatColumns="4" Width="90%">
                        <asp:ListItem Value="00" Text="全部報修單" Selected="True" />
                        <asp:ListItem Value="01" Text="未派工" />
                        <asp:ListItem Value="02" Text="施工中" />
                        <asp:ListItem Value="03" Text="已完工" />
                    </asp:RadioButtonList>
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbKeyword_Search" runat="server" CssClass="text-Right-Blue" Text="關鍵字：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eKeyword_Search" runat="server" CssClass="text-Left-Black" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
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
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridFixRequestList" runat="server" AllowPaging="True" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataKeyNames="sheetno" DataSourceID="sdsFixRequestList" GridLines="None" PageSize="5" Width="100%">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="sheetno" HeaderText="報修單號" ReadOnly="True" SortExpression="sheetno" />
                <asp:BoundField DataField="day" HeaderText="報修日期" SortExpression="day" DataFormatString="{0:d}" />
                <asp:BoundField DataField="BuMan" HeaderText="報修人" ReadOnly="True" SortExpression="BuMan" />
                <asp:BoundField DataField="FixType" HeaderText="需求類別" ReadOnly="True" SortExpression="FixType" />
                <asp:BoundField DataField="FixNote" HeaderText="報修內容" ReadOnly="True" SortExpression="FixNote" />
                <asp:BoundField DataField="FixRemark" HeaderText="內容說明" SortExpression="FixRemark" />
                <asp:BoundField DataField="Handler" HeaderText="處理人員" ReadOnly="True" SortExpression="Handler" />
                <asp:BoundField DataField="AssignDate" HeaderText="派工日期" SortExpression="AssignDate" DataFormatString="{0:d}" />
                <asp:BoundField DataField="Disposal" HeaderText="處理進度" ReadOnly="True" SortExpression="Disposal" />
                <asp:BoundField DataField="FixFinishDate" HeaderText="完工日期" SortExpression="FixFinishDate" DataFormatString="{0:d}" />
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
        <asp:FormView ID="fvFixRequestDetail" runat="server" DataKeyNames="sheetno" DataSourceID="sdsFixRequestDetail" Width="100%" OnDataBound="fvFixRequestDetail_DataBound">
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Black" CommandName="Update" Text="確定" Width="120px" />
                <asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbSheetNo_Edit" runat="server" CssClass="text-Right-Blue" Text="報修單號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eSheetNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("sheetno") %>' Width="90%" />
                                    <asp:Label ID="eDisposal_Edit" runat="server" Text='<%# Eval("Disposal") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbDay_Edit" runat="server" CssClass="text-Right-Blue" Text="報修日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eDay_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("day","{0:yyyy/MM/dd}") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="報修人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="eBuMan_Edit" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                                    <asp:Label ID="eBuMan_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-High ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbFixNote_Edit" runat="server" CssClass="text-Right-Blue" Text="報修內容：" Width="90%" />
                                </td>
                                <td class="MultiLine-High ColBorder ColWidth-6Col" colspan="7">
                                    <asp:Label ID="eFixType_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FixType") %>' Height="25px" Width="150px" />
                                    <asp:Label ID="eFixNote_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FixNote") %>' Height="25px" Width="300px" />
                                    <br />
                                    <asp:TextBox ID="eFixRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("FixRemark") %>' Height="60px" Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbAssignDate_Edit" runat="server" CssClass="text-Right-Blue" Text="派工日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eAssignDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbHandler_Edit" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:DropDownList ID="ddlHandler_Edit" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True"
                                        DataSourceID="sdsEmployee_Dep09" DataTextField="NAME" DataValueField="EMPNO" OnTextChanged="ddlHandler_Edit_TextChanged">
                                    </asp:DropDownList>
                                    <asp:Label ID="eHandler_C_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Handler_C") %>' Visible="false" />
                                    <asp:Label ID="eHandler_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Bind("Handler") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbFinishDate_Edit" runat="server" CssClass="text-Right-Blue" Text="完工日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-6Col">
                                    <asp:TextBox ID="eFinishDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FixFinishDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbChiefNote_Edit" runat="server" CssClass="text-Right-Blue" Text="主管交辦：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eChiefNote_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Bind("ChiefNote") %>' Height="95%" Width="97%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-6Col">
                                    <asp:Label ID="lbWorkReport_Edit" runat="server" CssClass="text-Right-Blue" Text="處置情況：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="2">
                                    <asp:TextBox ID="eWorkReport_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Bind("WorkReport") %>' Height="95%" Width="97%" />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <ItemTemplate>
                <asp:Button ID="bbModify" runat="server" CssClass="button-Blue" CausesValidation="False" CommandName="Edit" Text="編輯報修單" Width="120px" />
                <asp:Button ID="bbDel" runat="server" CssClass="button-Red" CausesValidation="False" OnClick="bbDel_Click" Text="刪除報修單" Width="120px" />
                <asp:Button ID="bbPrint" runat="server" CssClass="button-Black" CausesValidation="False" OnClick="bbPrint_Click" Text="列印報修單" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbSheetNo_List" runat="server" CssClass="text-Right-Blue" Text="報修單號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eSheetNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("sheetno") %>' Width="90%" />
                            <asp:Label ID="eDisposal_List" runat="server" Text='<%# Eval("Disposal") %>' Visible="true" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbDay_List" runat="server" CssClass="text-Right-Blue" Text="報修日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eDay_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("day","{0:yyyy/MM/dd}") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="報修人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eBuMan_List" runat="server" Text='<%# Eval("BuMan") %>' Visible="false" />
                            <asp:Label ID="eBuMan_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan_C") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-High ColBorder ColWidth-6Col">
                            <asp:Label ID="lbFixNote_List" runat="server" CssClass="text-Right-Blue" Text="報修內容：" Width="90%" />
                        </td>
                        <td class="MultiLine-High ColBorder ColWidth-6Col" colspan="7">
                            <asp:Label ID="eFixType_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FixType") %>' Height="25px" Width="150px" />
                            <asp:Label ID="eFixNote_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FixNote") %>' Height="25px" Width="300px" />
                            <br />
                            <asp:TextBox ID="eFixRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("FixRemark") %>' Height="60px" Width="97%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbAssignDate_List" runat="server" CssClass="text-Right-Blue" Text="派工日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eAssignDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("AssignDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbHandler_List" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eHandler_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Handler_C") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbFinishDate_List" runat="server" CssClass="text-Right-Blue" Text="完工日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eFinishDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FixFinishDate","{0:yyyy/MM/dd}") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbChiefNote_List" runat="server" CssClass="text-Right-Blue" Text="主管交辦：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="2">
                            <asp:TextBox ID="eChiefNote_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("ChiefNote") %>' Height="95%" Width="97%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col">
                            <asp:Label ID="lbWorkReport_List" runat="server" CssClass="text-Right-Blue" Text="處置情況：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-6Col" colspan="2">
                            <asp:TextBox ID="eWorkReport_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("WorkReport") %>' Height="95%" Width="97%" />
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
        <asp:Button ID="bbCloseReport" runat="server" CssClass="button-Red" Text="關閉報表" OnClick="bbCloseReport_Click" Width="100px" />
        <rsweb:ReportViewer ID="rvPrint" runat="server" BackColor="" ClientIDMode="AutoID" HighlightBackgroundColor=""
            InternalBorderColor="204, 204, 204" InternalBorderStyle="Solid" InternalBorderWidth="1px"
            LinkActiveColor="" LinkActiveHoverColor="" LinkDisabledColor=""
            PrimaryButtonBackgroundColor="" PrimaryButtonForegroundColor="" PrimaryButtonHoverBackgroundColor="" PrimaryButtonHoverForegroundColor=""
            SecondaryButtonBackgroundColor="" SecondaryButtonForegroundColor="" SecondaryButtonHoverBackgroundColor="" SecondaryButtonHoverForegroundColor="" SplitterBackColor=""
            ToolbarDividerColor="" ToolbarForegroundColor="" ToolbarForegroundDisabledColor="" ToolbarHoverBackgroundColor="" ToolbarHoverForegroundColor=""
            ToolBarItemBorderColor="" ToolBarItemBorderStyle="Solid" ToolBarItemBorderWidth="1px" ToolBarItemHoverBackColor="" ToolBarItemPressedBorderColor="51, 102, 153"
            ToolBarItemPressedBorderStyle="Solid" ToolBarItemPressedBorderWidth="1px" ToolBarItemPressedHoverBackColor="153, 187, 226" Height="600px" Width="100%">
            <LocalReport ReportPath="Report\FixSheetP.rdlc">
                <DataSources>
                </DataSources>
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsFixRequestList" runat="server"
        ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT sheetno, day, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = s.buman)) AS BuMan, CASE WHEN isnull(FixType , '01') = '01' THEN '軟體' WHEN isnull(FixType , '01') = '02' THEN '硬體' ELSE '資訊評估' END AS FixType, CASE WHEN isnull(aboutReport , 'X') = 'V' THEN '列印報表' ELSE '' END + CASE WHEN isnull(aboutDataModify , 'X') = 'V' THEN '資料修改' ELSE '' END + CASE WHEN isnull(aboutDesign , 'X') = 'V' THEN '設計資料或表單' ELSE '' END + CASE WHEN isnull(aboutHardDriver , 'X') = 'V' THEN '電腦故障' ELSE '' END + CASE WHEN isnull(aboutSetting , 'X') = 'V' THEN '安裝或設定' ELSE '' END + CASE WHEN isnull(aboutSurver , 'X') = 'V' THEN '偵測或補發' ELSE '' END + CASE WHEN isnull(aboutPurchase , 'X') = 'V' THEN '購置' ELSE '' END + CASE WHEN isnull(aboutOthers , 'X') = 'V' THEN '其他' ELSE '' END AS FixNote, FixRemark, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = s.Handler)) AS Handler, AssignDate, CASE WHEN isnull(Disposal , '01') = '01' THEN '未派工' WHEN isnull(Disposal , '01') = '02' THEN '施工中' ELSE '已完工' END AS Disposal, FixFinishDate FROM SHEETA AS s WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsFixRequestDetail" runat="server"
        ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT sheetno, day, Disposal, buman, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = s.buman)) AS BuMan_C, CASE WHEN isnull(FixType , '01') = '01' THEN '軟體' WHEN isnull(FixType , '01') = '02' THEN '硬體' ELSE '資訊評估' END AS FixType, CASE WHEN isnull(aboutReport , 'X') = 'V' THEN '列印報表' ELSE '' END + CASE WHEN isnull(aboutDataModify , 'X') = 'V' THEN '資料修改' ELSE '' END + CASE WHEN isnull(aboutDesign , 'X') = 'V' THEN '設計資料或表單' ELSE '' END + CASE WHEN isnull(aboutHardDriver , 'X') = 'V' THEN '電腦故障' ELSE '' END + CASE WHEN isnull(aboutSetting , 'X') = 'V' THEN '安裝或設定' ELSE '' END + CASE WHEN isnull(aboutSurver , 'X') = 'V' THEN '偵測或補發' ELSE '' END + CASE WHEN isnull(aboutPurchase , 'X') = 'V' THEN '購置' ELSE '' END + CASE WHEN isnull(aboutOthers , 'X') = 'V' THEN '其他' ELSE '' END AS FixNote, FixRemark, AssignDate, Handler, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = s.Handler)) AS Handler_C, ChiefNote, WorkReport, FixFinishDate FROM SHEETA AS s WHERE (sheetno = @sheetno)"
        UpdateCommand="UPDATE SHEETA SET AssignDate = @AssignDate, Handler = @Handler, ChiefNote = @ChiefNote, FixFinishDate = @FixFinishDate, WorkReport = @WorkReport, Disposal = @Disposal WHERE (sheetno = @SheetNo)"
        OnUpdating="sdsFixRequestDetail_Updating">
        <SelectParameters>
            <asp:ControlParameter ControlID="gridFixRequestList" Name="sheetno" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="AssignDate" />
            <asp:Parameter Name="Handler" />
            <asp:Parameter Name="ChiefNote" />
            <asp:Parameter Name="FixFinishDate" />
            <asp:Parameter Name="WorkReport" />
            <asp:Parameter Name="SheetNo" />
            <asp:Parameter Name="Disposal" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsEmployee_Dep09" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT CAST('' AS varchar) AS EmpNo, CAST('' AS varchar) AS Name UNION ALL SELECT EMPNO, NAME FROM EMPLOYEE WHERE (DEPNO = '09') AND (LEAVEDAY IS NULL)"></asp:SqlDataSource>
</asp:Content>
