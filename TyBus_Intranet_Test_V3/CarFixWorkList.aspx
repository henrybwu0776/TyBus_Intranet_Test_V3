<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CarFixWorkList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CarFixWorkList" %>

<asp:Content ID="CarFixWorkListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">修車廠工作單資料匯出</a>
    </div>
    <br />
    <asp:panel id="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlDepNo_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="True" OnSelectedIndexChanged="ddlDepNo_Search_SelectedIndexChanged" Width="90%" DataSourceID="sdsDepNoList" DataTextField="DepName" DataValueField="DEPNO"></asp:DropDownList>
                    <asp:TextBox ID="eDepNo_Search" runat="server" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbFixDateIn_Search" runat="server" CssClass="text-Right-Blue" Text="進廠日期：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eFixYear_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:TextBox ID="eFixMonth_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text=" 月" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="4">
                    <asp:RadioButtonList ID="rbOutputMode" runat="server" RepeatColumns="3" Width="90%">
                        <asp:ListItem Text="完整資料" Value="0" Selected="True" />
                        <asp:ListItem Text="工作項目與標準工時" Value="1" />
                        <asp:ListItem Text="工作人時" Value="2" />
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExcel" runat="server" CssClass="button-Black" Text="匯出 EXCEL" OnClick="bbExcel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="離開" OnClick="bbClose_Click" Width="90%" />
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
    </asp:panel>
    <asp:sqldatasource id="sdsDepNoList" runat="server" connectionstring="<%$ ConnectionStrings:connERPSQL %>" providername="<%$ ConnectionStrings:connERPSQL.ProviderName %>" selectcommand="SELECT DISTINCT DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = f.DEPNO)) AS DepName FROM FixworkA AS f"></asp:sqldatasource>
</asp:Content>
