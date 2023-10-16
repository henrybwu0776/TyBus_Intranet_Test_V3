<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="EmpAltersReport.aspx.cs" Inherits="TyBus_Intranet_Test_V3.EmpAltersReport" %>

<asp:Content ID="EmpAltersReportForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">人事異動單資料匯出</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbReportName_Search" runat="server" CssClass="text-Right-Blue" Text="報表格式：" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:DropDownList ID="ddlReportName_Search" runat="server" CssClass="text-Left-Blue" Width="95%">
                        <asp:ListItem Value="A000" Text="" Selected="True" />
                        <asp:ListItem Value="A001" Text="進用送審單" />
                        <asp:ListItem Value="A002" Text="僱用送審單" />
                        <asp:ListItem Value="A003" Text="調任送審單" />
                        <asp:ListItem Value="A004" Text="解聘送審單" />
                        <asp:ListItem Value="A005" Text="解聘送審單_審核" />
                        <asp:ListItem Value="A006" Text="聘任送審單" />
                        <asp:ListItem Value="A007" Text="退休送審單" />
                        <asp:ListItem Value="A008" Text="解聘送審單_2" />
                        <asp:ListItem Value="A009" Text="當然解聘送審單" />
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbAltersNo_Search" runat="server" CssClass="text-Right-Blue" Text="異動序號：" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eAltersNo_S_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="titleText-Black" Text="～" />
                    <asp:TextBox ID="eAltersNo_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbEmpNo_Search" runat="server" CssClass="text-Right-Blue" Text="員工編號：" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eEmpNo_Search" runat="server" CssClass="text-Left-Black" Width="35%" OnTextChanged="eEmpNo_Search_TextChanged" AutoPostBack="true" />
                    <asp:Label ID="eEmpName_Search" runat="server" CssClass="text-Left-Black" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="所屬部門：" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eDepNo_S_Search" runat="server" CssClass="text-Left-Black" Width="10%" OnTextChanged="eDepNo_S_Search_TextChanged" AutoPostBack="true" />
                    <asp:Label ID="eDepName_S_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="titleText-Black" Text="～" />
                    <asp:TextBox ID="eDepNo_E_Search" runat="server" CssClass="text-Left-Black" Width="10%" OnTextChanged="eDepNo_E_Search_TextChanged" AutoPostBack="true" />
                    <asp:Label ID="eDepName_E_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col">
                    <asp:Label ID="lbInureDate_Search" runat="server" CssClass="text-Right-Blue" Text="生效日期：" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                    <asp:TextBox ID="eInureDate_S_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                    <asp:Label ID="lbSplit_3" runat="server" CssClass="titleText-Black" Text="～" />
                    <asp:TextBox ID="eInureDate_E_Search" runat="server" CssClass="text-Left-Black" Width="40%" />
                </td>
            </tr>
            <tr>
                <td class="MultiLine_High ColBorder ColWidth-6Col">
                    <asp:Label ID="lbAlterType_Search" runat="server" CssClass="text-Right-Blue" Text="異動類別：" Width="100%" />
                </td>
                <td class="MultiLine_High ColBorder ColWidth-6Col" colspan="3">
                    <asp:CheckBoxList ID="eAlterType_Search" runat="server" CssClass="text-Left-Blue" Width="100%" AutoPostBack="True"
                        DataSourceID="dsAlterType" DataTextField="ClassTxt" DataValueField="ClassNo" Height="100%" RepeatColumns="4" />
                </td>
                <td class="ColHeight ColWidth-6Col" />
                <td class="ColHeight ColWidth-6Col">
                    <asp:Button ID="bbOK" runat="server" CssClass="button-Blue" Text="匯出" Width="45%" OnClick="bbOK_Click" />
                    <asp:Button ID="bbCancel" runat="server" CssClass="button-Red" Text="結束" Width="45%" OnClick="bbCancel_Click" />
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
    <asp:SqlDataSource ID="dsAlterType" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" SelectCommand="select ClassNo, ClassTxt from DBDICB where FKey = '人事異動檔      ALTERS          altertype' order by ClassNo"></asp:SqlDataSource>
</asp:Content>
