<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DriverWorkHoursDays.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DriverWorkHoursDays" %>
<asp:Content ID="DriverWorkHoursDaysForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">駕駛員時數統計查詢</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalYM_Search" runat="server" CssClass="text-Right-Blue" Text="查詢年月：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eCalYear_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbCalYear_Search" runat="server" CssClass="text-Left-Black" Text=" 年" Width="5%" />
                    <asp:TextBox ID="eCalMonth_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbCalMonth_Search" runat="server" CssClass="text-Left-Black" Text=" 月" Width="5%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDriver_Search" runat="server" CssClass="text-Right-Blue" Text="駕駛員工號：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eDriverNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriverNo_Search_TextChanged" Width="40%" />
                    <asp:Label ID="eDriverName_Search" runat="server" CssClass="text-Left-Black" Width="55%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClear" runat="server" CssClass="button-Red" Text="清除" OnClick="bbClear_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-6Col" colspan="3">
                    <asp:RequiredFieldValidator ID="rfvDriver" runat="server" ErrorMessage="駕駛員工號不可空白" ControlToValidate="eDriverNo_Search" Font-Bold="True" ForeColor="Red" SetFocusOnError="True"></asp:RequiredFieldValidator>
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
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel" Visible="false">
        <asp:FormView ID="fvDriverWorkHoursData" runat="server" DataSourceID="sdsDriverWorkHoursData" Width="100%">
            <EmptyDataTemplate>
                <asp:Label ID="EmptyText" runat="server" CssClass="Title1" Text="查無資料！" />
            </EmptyDataTemplate>
            <ItemTemplate>
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                            <asp:Label ID="lbWarningMSG" runat="server" CssClass="errorMessageText" Text="本報表資料直接取自行車憑單，僅供即時資料查詢" Width="90%" />
                            <br />
                            <asp:Label ID="lbWarningMSG_1" runat="server" CssClass="errorMessageText" Text="資料內容會隨行車憑單修改而變動" Width="90%" />
                            <br />
                            <asp:Label ID="lbWarningMSG_2" runat="server" CssClass="errorMessageText" Text="正確資料仍以主計單位每月完成關帳結算後為準" Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriver" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDriveNo" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver") %>' Width="35%" />
                            <asp:Label ID="eDriverName" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                            <br />
                            <asp:Label ID="eIsOverSeven" runat="server" Text="★ 超過七日未休" ForeColor="Red" Width="90%" Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRentNumber" runat="server" CssClass="text-Right-Blue" Text="趟次：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eRentNumber" runat="server" CssClass="text-Left-Black" Text='<%# Eval("RentNumber") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbTotalHR" runat="server" CssClass="text-Right-Blue" Text="總時數：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eTotalHR" runat="server" CssClass="text-Left-Black" Text='<%# Eval("TotalHR") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="Title1 ColBorder ColWidth-8Col" colspan="7">
                            <asp:Label ID="lbSubTitle" runat="server" Text="每日時數" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay1" runat="server" CssClass="textBlue" Text="1" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay2" runat="server" CssClass="textBlue" Text="2" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay3" runat="server" CssClass="textBlue" Text="3" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay4" runat="server" CssClass="textBlue" Text="4" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay5" runat="server" CssClass="textBlue" Text="5" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay6" runat="server" CssClass="textBlue" Text="6" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay7" runat="server" CssClass="textBlue" Text="7" Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay1" runat="server" CssClass="textBold" Text='<%# Eval("Hour01") %>' Width="55%" />
                            <asp:Label ID="eNoWork1" runat="server" CssClass="textRed" Text='<%# Eval("WorkState01") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay2" runat="server" CssClass="textBold" Text='<%# Eval("Hour02") %>' Width="55%" />
                            <asp:Label ID="eNoWork2" runat="server" CssClass="textRed" Text='<%# Eval("WorkState02") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay3" runat="server" CssClass="textBold" Text='<%# Eval("Hour03") %>' Width="55%" />
                            <asp:Label ID="eNoWork3" runat="server" CssClass="textRed" Text='<%# Eval("WorkState03") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay4" runat="server" CssClass="textBold" Text='<%# Eval("Hour04") %>' Width="55%" />
                            <asp:Label ID="eNoWork4" runat="server" CssClass="textRed" Text='<%# Eval("WorkState04") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay5" runat="server" CssClass="textBold" Text='<%# Eval("Hour05") %>' Width="55%" />
                            <asp:Label ID="eNoWork5" runat="server" CssClass="textRed" Text='<%# Eval("WorkState05") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay6" runat="server" CssClass="textBold" Text='<%# Eval("Hour06") %>' Width="55%" />
                            <asp:Label ID="eNoWork6" runat="server" CssClass="textRed" Text='<%# Eval("WorkState06") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay7" runat="server" CssClass="textBold" Text='<%# Eval("Hour07") %>' Width="55%" />
                            <asp:Label ID="eNoWork7" runat="server" CssClass="textRed" Text='<%# Eval("WorkState07") %>' Width="30%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay8" runat="server" CssClass="textBlue" Text="8" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay9" runat="server" CssClass="textBlue" Text="9" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay10" runat="server" CssClass="textBlue" Text="10" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay11" runat="server" CssClass="textBlue" Text="11" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay12" runat="server" CssClass="textBlue" Text="12" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay13" runat="server" CssClass="textBlue" Text="13" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay14" runat="server" CssClass="textBlue" Text="14" Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay8" runat="server" CssClass="textBold" Text='<%# Eval("Hour08") %>' Width="55%" />
                            <asp:Label ID="eNoWork8" runat="server" CssClass="textRed" Text='<%# Eval("WorkState08") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay9" runat="server" CssClass="textBold" Text='<%# Eval("Hour09") %>' Width="55%" />
                            <asp:Label ID="eNoWork9" runat="server" CssClass="textRed" Text='<%# Eval("WorkState09") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay10" runat="server" CssClass="textBold" Text='<%# Eval("Hour10") %>' Width="55%" />
                            <asp:Label ID="eNoWork10" runat="server" CssClass="textRed" Text='<%# Eval("WorkState10") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay11" runat="server" CssClass="textBold" Text='<%# Eval("Hour11") %>' Width="55%" />
                            <asp:Label ID="eNoWork11" runat="server" CssClass="textRed" Text='<%# Eval("WorkState11") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay12" runat="server" CssClass="textBold" Text='<%# Eval("Hour12") %>' Width="55%" />
                            <asp:Label ID="eNoWork12" runat="server" CssClass="textRed" Text='<%# Eval("WorkState12") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay13" runat="server" CssClass="textBold" Text='<%# Eval("Hour13") %>' Width="55%" />
                            <asp:Label ID="eNoWork13" runat="server" CssClass="textRed" Text='<%# Eval("WorkState13") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay14" runat="server" CssClass="textBold" Text='<%# Eval("Hour14") %>' Width="55%" />
                            <asp:Label ID="eNoWork14" runat="server" CssClass="textRed" Text='<%# Eval("WorkState14") %>' Width="30%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay15" runat="server" CssClass="textBlue" Text="15" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay16" runat="server" CssClass="textBlue" Text="16" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay17" runat="server" CssClass="textBlue" Text="17" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay18" runat="server" CssClass="textBlue" Text="18" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay19" runat="server" CssClass="textBlue" Text="19" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay20" runat="server" CssClass="textBlue" Text="20" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay21" runat="server" CssClass="textBlue" Text="21" Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay15" runat="server" CssClass="textBold" Text='<%# Eval("Hour15") %>' Width="55%" />
                            <asp:Label ID="eNoWork15" runat="server" CssClass="textRed" Text='<%# Eval("WorkState15") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay16" runat="server" CssClass="textBold" Text='<%# Eval("Hour16") %>' Width="55%" />
                            <asp:Label ID="eNoWork16" runat="server" CssClass="textRed" Text='<%# Eval("WorkState16") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay17" runat="server" CssClass="textBold" Text='<%# Eval("Hour17") %>' Width="55%" />
                            <asp:Label ID="eNoWork17" runat="server" CssClass="textRed" Text='<%# Eval("WorkState17") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay18" runat="server" CssClass="textBold" Text='<%# Eval("Hour18") %>' Width="55%" />
                            <asp:Label ID="eNoWork18" runat="server" CssClass="textRed" Text='<%# Eval("WorkState18") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay19" runat="server" CssClass="textBold" Text='<%# Eval("Hour19") %>' Width="55%" />
                            <asp:Label ID="eNoWork19" runat="server" CssClass="textRed" Text='<%# Eval("WorkState19") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay20" runat="server" CssClass="textBold" Text='<%# Eval("Hour20") %>' Width="55%" />
                            <asp:Label ID="eNoWork20" runat="server" CssClass="textRed" Text='<%# Eval("WorkState20") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay21" runat="server" CssClass="textBold" Text='<%# Eval("Hour21") %>' Width="55%" />
                            <asp:Label ID="eNoWork21" runat="server" CssClass="textRed" Text='<%# Eval("WorkState21") %>' Width="30%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay22" runat="server" CssClass="textBlue" Text="22" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay23" runat="server" CssClass="textBlue" Text="23" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay24" runat="server" CssClass="textBlue" Text="24" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay25" runat="server" CssClass="textBlue" Text="25" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay26" runat="server" CssClass="textBlue" Text="26" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay27" runat="server" CssClass="textBlue" Text="27" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay28" runat="server" CssClass="textBlue" Text="28" Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay22" runat="server" CssClass="textBold" Text='<%# Eval("Hour22") %>' Width="55%" />
                            <asp:Label ID="eNoWork22" runat="server" CssClass="textRed" Text='<%# Eval("WorkState22") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay23" runat="server" CssClass="textBold" Text='<%# Eval("Hour23") %>' Width="55%" />
                            <asp:Label ID="eNoWork23" runat="server" CssClass="textRed" Text='<%# Eval("WorkState23") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay24" runat="server" CssClass="textBold" Text='<%# Eval("Hour24") %>' Width="55%" />
                            <asp:Label ID="eNoWork24" runat="server" CssClass="textRed" Text='<%# Eval("WorkState24") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay25" runat="server" CssClass="textBold" Text='<%# Eval("Hour25") %>' Width="55%" />
                            <asp:Label ID="eNoWork25" runat="server" CssClass="textRed" Text='<%# Eval("WorkState25") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay26" runat="server" CssClass="textBold" Text='<%# Eval("Hour26") %>' Width="55%" />
                            <asp:Label ID="eNoWork26" runat="server" CssClass="textRed" Text='<%# Eval("WorkState26") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay27" runat="server" CssClass="textBold" Text='<%# Eval("Hour27") %>' Width="55%" />
                            <asp:Label ID="eNoWork27" runat="server" CssClass="textRed" Text='<%# Eval("WorkState27") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay28" runat="server" CssClass="textBold" Text='<%# Eval("Hour28") %>' Width="55%" />
                            <asp:Label ID="eNoWork28" runat="server" CssClass="textRed" Text='<%# Eval("WorkState28") %>' Width="30%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay29" runat="server" CssClass="textBlue" Text="29" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay30" runat="server" CssClass="textBlue" Text="30" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDay31" runat="server" CssClass="textBlue" Text="31" Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay29" runat="server" CssClass="textBold" Text='<%# Eval("Hour29") %>' Width="55%" />
                            <asp:Label ID="eNoWork29" runat="server" CssClass="textRed" Text='<%# Eval("WorkState29") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay30" runat="server" CssClass="textBold" Text='<%# Eval("Hour30") %>' Width="55%" />
                            <asp:Label ID="eNoWork30" runat="server" CssClass="textRed" Text='<%# Eval("WorkState30") %>' Width="30%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eDay31" runat="server" CssClass="textBold" Text='<%# Eval("Hour31") %>' Width="55%" />
                            <asp:Label ID="eNoWork31" runat="server" CssClass="textRed" Text='<%# Eval("WorkState31") %>' Width="30%" />
                        </td>
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsDriverWorkHoursData" runat="server" 
        ConnectionString="<%$ ConnectionStrings:connERPSQL %>" 
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" 
        SelectCommand="SELECT a.DRIVER, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.DRIVER)) AS DriverName, CAST(0 AS float) AS RentNumber, CAST(0 AS float) AS TotalHR, b.Hour01, b.Hour02, b.Hour03, b.Hour04, b.Hour05, b.Hour06, b.Hour07, b.Hour08, b.Hour09, b.Hour10, b.Hour11, b.Hour12, b.Hour13, b.Hour14, b.Hour15, b.Hour16, b.Hour17, b.Hour18, b.Hour19, b.Hour20, b.Hour21, b.Hour22, b.Hour23, b.Hour24, b.Hour25, b.Hour26, b.Hour27, b.Hour28, b.Hour29, b.Hour30, b.Hour31, LEFT (b.WORKSTATE01, 1) AS WorkState01, LEFT (b.WORKSTATE02, 1) AS WorkState02, LEFT (b.WORKSTATE03, 1) AS WorkState03, LEFT (b.WORKSTATE04, 1) AS WorkState04, LEFT (b.WORKSTATE05, 1) AS WorkState05, LEFT (b.WORKSTATE06, 1) AS WorkState06, LEFT (b.WORKSTATE07, 1) AS WorkState07, LEFT (b.WORKSTATE08, 1) AS WorkState08, LEFT (b.WORKSTATE09, 1) AS WorkState09, LEFT (b.WORKSTATE10, 1) AS WorkState10, LEFT (b.WORKSTATE11, 1) AS WorkState11, LEFT (b.WORKSTATE12, 1) AS WorkState12, LEFT (b.WORKSTATE13, 1) AS WorkState13, LEFT (b.WORKSTATE14, 1) AS WorkState14, LEFT (b.WORKSTATE15, 1) AS WorkState15, LEFT (b.WORKSTATE16, 1) AS WorkState16, LEFT (b.WORKSTATE17, 1) AS WorkState17, LEFT (b.WORKSTATE18, 1) AS WorkState18, LEFT (b.WORKSTATE19, 1) AS WorkState19, LEFT (b.WORKSTATE20, 1) AS WorkState20, LEFT (b.WORKSTATE21, 1) AS WorkState21, LEFT (b.WORKSTATE22, 1) AS WorkState22, LEFT (b.WORKSTATE23, 1) AS WorkState23, LEFT (b.WORKSTATE24, 1) AS WorkState24, LEFT (b.WORKSTATE25, 1) AS WorkState25, LEFT (b.WORKSTATE26, 1) AS WorkState26, LEFT (b.WORKSTATE27, 1) AS WorkState27, LEFT (b.WORKSTATE28, 1) AS WorkState28, LEFT (b.WORKSTATE29, 1) AS WorkState29, LEFT (b.WORKSTATE30, 1) AS WorkState30, LEFT (b.WORKSTATE31, 1) AS WorkState31 FROM RUNSHEETA AS a LEFT OUTER JOIN DriverOverDuty AS b ON b.OverYM = CAST(YEAR(a.BUDATE) AS varchar(4)) + RIGHT ('00' + CAST(MONTH(a.BUDATE) AS varchar(2)), 2) AND b.EMPNO = a.DRIVER WHERE (1 &lt;&gt; 1)">
    </asp:SqlDataSource>
</asp:Content>
