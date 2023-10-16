<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="FixRequestMain.aspx.cs" Inherits="TyBus_Intranet_Test_V3.FixRequestMain" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="FixRequestMain" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">線上報修暨進度查詢</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbBuMan" runat="server" CssClass="text-Right-Blue" Text="申請人：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eBuMan_Search" runat="server" CssClass="text-Left-Black" Width="40%" OnTextChanged="eBuMan_Search_TextChanged" AutoPostBack="True" />
                    <asp:Label ID="eBuManName_Search" runat="server" CssClass="text-Left-Black" Width="50%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbBuDate_Search" runat="server" CssClass="text-Right-Blue" Text="申請日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eBudate_S_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="～" Width="5%" />
                    <asp:TextBox ID="eBudate_E_Search" runat="server" CssClass="text-Left-Black" Width="35%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
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
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:GridView ID="gridShowData" runat="server" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" GridLines="None" AutoGenerateColumns="False" DataKeyNames="sheetno" DataSourceID="sdsFixRequestList"
            Width="100%" AllowPaging="True" PageSize="5">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="sheetno" HeaderText="報修單號" ReadOnly="True" SortExpression="sheetno" />
                <asp:BoundField DataField="day" DataFormatString="{0:d}" HeaderText="報修日期" SortExpression="day" />
                <asp:BoundField DataField="buman" HeaderText="buman" SortExpression="buman" Visible="False" />
                <asp:BoundField DataField="BuManName" HeaderText="報修人" ReadOnly="True" SortExpression="BuManName" />
                <asp:BoundField DataField="FixType" HeaderText="需求類別" ReadOnly="True" SortExpression="FixType" />
                <asp:BoundField DataField="FixNote" HeaderText="報修內容" ReadOnly="True" SortExpression="FixNote" />
                <asp:BoundField DataField="FixRemark" HeaderText="內容說明" SortExpression="FixRemark" />
                <asp:BoundField DataField="Handler" HeaderText="處理人員" ReadOnly="True" SortExpression="Handler" />
                <asp:BoundField DataField="AssignDate" DataFormatString="{0:d}" HeaderText="派工日期" SortExpression="AssignDate" />
                <asp:BoundField DataField="Disposal" HeaderText="處理進度" ReadOnly="True" SortExpression="Disposal" />
                <asp:BoundField DataField="FixFinishDate" DataFormatString="{0:d}" HeaderText="完工日期" SortExpression="FixFinishDate" />
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
    <asp:Panel ID="plDetailData" runat="server" CssClass="ShowPanel-Detail">
        <asp:FormView ID="fvFixRequest" runat="server" DataKeyNames="sheetno" DataSourceID="sdsFixRequestMain" Width="100%" OnDataBound="fvFixRequest_DataBound">
            <EmptyDataTemplate>
                <asp:Button ID="bbCreateFix_Empty" runat="server" CssClass="button-Blue" CausesValidation="False" CommandName="New" Text="申請報修" Width="100px" />
            </EmptyDataTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CssClass="button-Blue" CausesValidation="True" OnClick="bbOK_INS_Click" Text="確定報修" Width="100px" OnClientClick="this.disabled=true;" UseSubmitBehavior="False" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CssClass="button-Red" CausesValidation="False" CommandName="Cancel" Text="取消" Width="100px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbSheetNo_Insert" runat="server" CssClass="text-Right-Blue" Text="報修單號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eSheetNo_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("sheetno") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFixType_Insert" runat="server" CssClass="text-Right-Blue" Text="需求類別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlFixType_Insert" runat="server" CssClass="text-Left-Black" OnSelectedIndexChanged="ddlFixType_Insert_SelectedIndexChanged" Width="90%">
                                        <asp:ListItem Value="" Text="" Selected="True" />
                                        <asp:ListItem Value="01" Text="軟體" />
                                        <asp:ListItem Value="02" Text="硬體" />
                                        <asp:ListItem Value="03" Text="資訊評估" />
                                    </asp:DropDownList>
                                    <asp:Label ID="eFixType_Insert" runat="server" Text='<%# Bind("FixType") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDisposal_Insert" runat="server" CssClass="text-Right-Blue" Text="處理進度：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eDisposal_Insert" runat="server" CssClass="text-Left-Black" Text="未派工" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbHandler_Insert" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eHandler_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("HandlerName") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDay_Insert" runat="server" CssClass="text-Right-Blue" Text="報修日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eDay_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("day", "{0:yyyy/MM/dd}") %>' />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepartment" runat="server" CssClass="text-Right-Blue" Text="報修單位：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("depno") %>' Width="35%" />
                                    <asp:Label ID="eDepName_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_Insert" runat="server" CssClass="text-Right-Blue" Text="報修人：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eBuMan_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Bind("buman") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_Insert" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eAboutReport_Insert" runat="server" CssClass="text-Left-Black" Enabled="true" Text="列印報表" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eAboutDataModify_Insert" runat="server" CssClass="text-Left-Black" Enabled="true" Text="修改資料" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eAboutDesign_Insert" runat="server" CssClass="text-Left-Black" Enabled="true" Text="設計資料或表單" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eAboutHardDriver_Insert" runat="server" CssClass="text-Left-Black" Enabled="true" Text="電腦故障" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eAboutSetting_Insert" runat="server" CssClass="text-Left-Black" Enabled="true" Text="安裝或設定" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eAboutSurver_Insert" runat="server" CssClass="text-Left-Black" Enabled="true" Text="偵測或補發" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eAboutPurchase_Insert" runat="server" CssClass="text-Left-Black" Enabled="true" Text="購置" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:CheckBox ID="eAboutOther_Insert" runat="server" CssClass="text-Left-Black" Enabled="true" Text="其他" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbFixRemark_Insert" runat="server" CssClass="text-Right-Blue" Text="報修內容：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                                    <asp:TextBox ID="eFixRemark_Insert" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="true" Text='<%# Bind("FixRemark") %>' Height="95%" Width="97%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-8Col">
                                    <asp:Label ID="eSheetType_Insert" runat="server" Text='<%# Bind("SheetType") %>' Visible="false" />
                                </td>
                                <td class="ColWidth-8Col">
                                    <asp:Label ID="eSheetClass_Insert" runat="server" Text='<%# Bind("SheetClass") %>' Visible="false" />
                                </td>
                                <td class="ColWidth-8Col">
                                    <asp:Label ID="eMode_Insert" runat="server" Text='<%# Bind("Mode") %>' Visible="false" />
                                </td>
                                <td class="ColWidth-8Col">
                                    <asp:Label ID="eUnitChief_Insert" runat="server" Text='<%# Bind("UnitChief") %>' Visible="false" />
                                </td>
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
                <asp:Button ID="bbCreateFix_List" runat="server" CausesValidation="False" CssClass="button-Blue" CommandName="New" Text="申請報修" Width="100px" />
                <asp:Button ID="bbPrintFixQuery_List" runat="server" CausesValidation="false" CssClass="button-Red" Text="列印報修單" OnClick="bbPrintFixQuery_List_Click" Width="100px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbSheetNo_List" runat="server" CssClass="text-Right-Blue" Text="報修單號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eSheetNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("sheetno") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbFixType_List" runat="server" CssClass="text-Right-Blue" Text="需求類別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eFixType_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("FixType_C") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDisposal_List" runat="server" CssClass="text-Right-Blue" Text="處理進度：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDisposal_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Disposal_C") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbHandler_List" runat="server" CssClass="text-Right-Blue" Text="處理人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eHandler_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("HandlerName") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay_List" runat="server" CssClass="text-Right-Blue" Text="報修日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("day", "{0:yyyy/MM/dd}") %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepartment" runat="server" CssClass="text-Right-Blue" Text="報修單位：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("depno") %>' Width="35%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="55%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="報修人：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("buman") %>' Width="35%" />
                            <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eAboutReport_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="列印報表" Checked='<%# Eval("aboutReport").ToString().Trim() == "V" %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eAboutDataModify_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="修改資料" Checked='<%# Eval("aboutDataModify").ToString().Trim() == "V" %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eAboutDesign_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="設計資料或表單" Checked='<%# Eval("aboutDesign").ToString().Trim() == "V" %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eAboutHardDriver_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="電腦故障" Checked='<%# Eval("aboutHardDriver").ToString().Trim() == "V" %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eAboutSetting_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="安裝或設定" Checked='<%# Eval("aboutSetting").ToString().Trim() == "V" %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eAboutSurver_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="偵測或補發" Checked='<%# Eval("aboutSurver").ToString().Trim() == "V" %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eAboutPurchase_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="購置" Checked='<%# Eval("aboutPurchase").ToString().Trim() == "V" %>' />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:CheckBox ID="eAboutOther_List" runat="server" CssClass="text-Left-Black" Enabled="false" Text="其他" Checked='<%# Eval("aboutOthers").ToString().Trim() == "V" %>' />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbFixRemark_List" runat="server" CssClass="text-Right-Blue" Text="報修內容：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="7">
                            <asp:TextBox ID="eFixRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("FixRemark") %>' Height="95%" Width="97%" />
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
    <asp:SqlDataSource ID="sdsFixRequestList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT sheetno, day, buman, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = s.buman)) AS BuManName, CASE WHEN isnull(s.FixType , '01') = '01' THEN '軟體' WHEN isnull(s.FixType , '01') = '02' THEN '硬體' ELSE '資訊評估' END AS FixType, CASE WHEN isnull(aboutReport , 'X') = 'V' THEN '列印報表' ELSE '' END + CASE WHEN isnull(aboutDataModify , 'X') = 'V' THEN '資料修改' ELSE '' END + CASE WHEN isnull(aboutDesign , 'X') = 'V' THEN '設計資料或表單' ELSE '' END + CASE WHEN isnull(aboutHardDriver , 'X') = 'V' THEN '電腦故障' ELSE '' END + CASE WHEN isnull(aboutSetting , 'X') = 'V' THEN '安裝或設定' ELSE '' END + CASE WHEN isnull(aboutSurver , 'X') = 'V' THEN '偵測或補發' ELSE '' END + CASE WHEN isnull(aboutPurchase , 'X') = 'V' THEN '購置' ELSE '' END + CASE WHEN isnull(aboutOthers , 'X') = 'V' THEN '其他' ELSE '' END AS FixNote, FixRemark, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = s.Handler)) AS Handler, AssignDate, CASE WHEN isnull(Disposal , '01') = '01' THEN '未派工' WHEN isnull(Disposal , '01') = '02' THEN '施工中' ELSE '已完工' END AS Disposal, FixFinishDate FROM SHEETA AS s WHERE (sheettype = 'Z') AND (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsFixRequestMain" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        InsertCommand="INSERT INTO SHEETA(sheetno, sheettype, day, buman, sheetclass, mode, UnitChief, FixType, aboutReport, aboutDataModify, aboutDesign, aboutHardDriver, aboutSetting, aboutSurver, aboutPurchase, aboutOthers, FixRemark, Disposal, depno) VALUES (@SheetNo, @SheetType, @Day, @BuMan, @SheetClass, @Mode, @UnitChief, @FixType, @aboutReport, @aboutDataModify, @aboutDesign, @aboutHardDriver, @aboutSetting, @aboutSurver, @aboutPurchase, @aboutOthers, @FixRemark, @Disposal, @DepNo)"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT sheetno, depno, FixType, CASE WHEN isnull(s.FixType , '01') = '01' THEN '軟體' WHEN isnull(s.FixType , '01') = '02' THEN '硬體' ELSE '資訊評估' END AS FixType_C, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = s.depno)) AS DepName, day, buman, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = s.buman)) AS BuManName, sheettype, sheetclass, mode, UnitChief, aboutReport, aboutDataModify, aboutDesign, aboutHardDriver, aboutSetting, aboutSurver, aboutPurchase, aboutOthers, FixRemark, Disposal, CASE WHEN isnull(Disposal , '01') = '01' THEN '未派工' WHEN isnull(Disposal , '01') = '02' THEN '施工中' ELSE '已完工' END AS Disposal_C, Handler, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = s.Handler)) AS HandlerName FROM SHEETA AS s WHERE (sheetno = @SheetNo)">
        <InsertParameters>
            <asp:Parameter Name="SheetNo" />
            <asp:Parameter Name="SheetType" />
            <asp:Parameter Name="Day" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="SheetClass" />
            <asp:Parameter Name="Mode" />
            <asp:Parameter Name="UnitChief" />
            <asp:Parameter Name="FixType" />
            <asp:Parameter Name="aboutReport" />
            <asp:Parameter Name="aboutDataModify" />
            <asp:Parameter Name="aboutDesign" />
            <asp:Parameter Name="aboutHardDriver" />
            <asp:Parameter Name="aboutSetting" />
            <asp:Parameter Name="aboutSurver" />
            <asp:Parameter Name="aboutPurchase" />
            <asp:Parameter Name="aboutOthers" />
            <asp:Parameter Name="FixRemark" />
            <asp:Parameter Name="Disposal" />
            <asp:Parameter Name="DepNo" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridShowData" Name="SheetNo" PropertyName="SelectedValue" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:Panel ID="plPrint" runat="server" CssClass="PrintPanel" Visible="false">
        <asp:Button ID="HidePrint" runat="server" CssClass="button-Red" Text="關閉報表" Width="100px" OnClick="HidePrint_Click" />
        <rsweb:ReportViewer ID="rvPrint" runat="server" BackColor="" ClientIDMode="AutoID" HighlightBackgroundColor=""
            InternalBorderColor="204, 204, 204" InternalBorderStyle="Solid" InternalBorderWidth="1px"
            LinkActiveColor="" LinkActiveHoverColor="" LinkDisabledColor=""
            PrimaryButtonBackgroundColor="" PrimaryButtonForegroundColor="" PrimaryButtonHoverBackgroundColor="" PrimaryButtonHoverForegroundColor=""
            SecondaryButtonBackgroundColor="" SecondaryButtonForegroundColor="" SecondaryButtonHoverBackgroundColor="" SecondaryButtonHoverForegroundColor=""
            SplitterBackColor="" ToolbarDividerColor="" ToolbarForegroundColor="" ToolbarForegroundDisabledColor="" ToolbarHoverBackgroundColor="" ToolbarHoverForegroundColor=""
            ToolBarItemBorderColor="" ToolBarItemBorderStyle="Solid" ToolBarItemBorderWidth="1px" ToolBarItemHoverBackColor="" ToolBarItemPressedBorderColor="51, 102, 153"
            ToolBarItemPressedBorderStyle="Solid" ToolBarItemPressedBorderWidth="1px" ToolBarItemPressedHoverBackColor="153, 187, 226" Height="600px" Width="100%">
            <LocalReport ReportPath="Report\FixRequestP.rdlc">
            </LocalReport>
        </rsweb:ReportViewer>
    </asp:Panel>
</asp:Content>
