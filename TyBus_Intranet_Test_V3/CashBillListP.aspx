<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CashBillListP.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CashBillListP" %>

<asp:Content ID="CashBillListPForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">投現繳銷統計</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="3">
                    <asp:RadioButtonList ID="rbListMode" runat="server" CssClass="text-Left-Black" RepeatColumns="3" Width="97%">
                        <asp:ListItem Value="1" Text="月繳銷清冊_依駕駛員統計" Selected="True" />
                        <asp:ListItem Value="2" Text="月繳銷清冊_依車輛別統計" />
                    </asp:RadioButtonList>
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlDepNo" runat="server" CssClass="text-Left-Black" AutoPostBack="True" OnSelectedIndexChanged="ddlDepNo_SelectedIndexChanged"
                        DataSourceID="sdsDepNoList" DataTextField="DepName" DataValueField="DEPNO"></asp:DropDownList>
                    <br />
                    <asp:TextBox ID="eDepNo" runat="server" CssClass="text-Left-Black" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbMonthListDate" runat="server" CssClass="text-Right-Blue" Text="統計月份：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                    <asp:TextBox ID="eMonthList_Year" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text="年" Width="5%" />
                    <asp:TextBox ID="eMonthList_Month" runat="server" CssClass="text-Left-Black" Width="35%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text="月" Width="5%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
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
    <asp:SqlDataSource ID="sdsDepNoList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" 
        SelectCommand="SELECT DEPNO, NAME AS DepName FROM DEPARTMENT WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
</asp:Content>
