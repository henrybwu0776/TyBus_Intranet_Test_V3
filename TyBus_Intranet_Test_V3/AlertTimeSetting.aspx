<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="AlertTimeSetting.aspx.cs" Inherits="TyBus_Intranet_Test_V3.AlertTimeSetting" %>
<asp:Content ID="AlertTimeSettingForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">提示視窗時間設定</a>
    </div>
    <br />
    <asp:Panel ID="plShow" runat="server" CssClass="ShowPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-4Col">
                    <asp:Label ID="lbAlertTime01" runat="server" CssClass="titleText-S-Blue" Text="第一次顯示" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-4Col">
                    <asp:Label ID="lbAlertTime02" runat="server" CssClass="titleText-S-Blue" Text="第二次顯示" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-4Col">
                    <asp:Label ID="lbAlertTime03" runat="server" CssClass="titleText-S-Blue" Text="第三次顯示" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-4Col">
                    <asp:Label ID="lbAlertTime04" runat="server" CssClass="titleText-S-Blue" Text="第四次顯示" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-4Col">
                    <asp:TextBox ID="eAlertTime01" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True" />
                    <br>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ControlToValidate="eAlertTime01" ErrorMessage="RegularExpressionValidator" ValidationExpression="\d{2}:\d{2}" CssClass="text-Left-Red" SetFocusOnError="True">時間格式錯誤</asp:RegularExpressionValidator>
                </td>
                <td class="ColHeight ColBorder ColWidth-4Col">
                    <asp:TextBox ID="eAlertTime02" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True" />
                    <br>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator2" runat="server" ControlToValidate="eAlertTime02" ErrorMessage="RegularExpressionValidator" ValidationExpression="\d{2}:\d{2}" CssClass="text-Left-Red" SetFocusOnError="True">時間格式錯誤</asp:RegularExpressionValidator>
                </td>
                <td class="ColHeight ColBorder ColWidth-4Col">
                    <asp:TextBox ID="eAlertTime03" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True" />
                    <br>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator3" runat="server" ControlToValidate="eAlertTime03" ErrorMessage="RegularExpressionValidator" ValidationExpression="\d{2}:\d{2}" CssClass="text-Left-Red" SetFocusOnError="True">時間格式錯誤</asp:RegularExpressionValidator>
                </td>
                <td class="ColHeight ColBorder ColWidth-4Col">
                    <asp:TextBox ID="eAlertTime04" runat="server" CssClass="text-Left-Black" Width="90%" AutoPostBack="True" />
                    <br>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator4" runat="server" ControlToValidate="eAlertTime04" ErrorMessage="RegularExpressionValidator" ValidationExpression="\d{2}:\d{2}" CssClass="text-Left-Red" SetFocusOnError="True">時間格式錯誤</asp:RegularExpressionValidator>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-4Col">
                    <asp:Button ID="bbModify" runat="server" CssClass="button-Blue" OnClick="bbModify_Click" Text="修改" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-4Col">
                    <asp:Button ID="bbOK" runat="server" CssClass="button-Black" OnClick="bbOK_Click" Text="存檔" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-4Col">
                    <asp:Button ID="bbCancel" runat="server" CssClass="button-Red" OnClick="bbCancel_Click" Text="放棄" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-4Col">
                    <asp:Button ID="bbExit" runat="server" CssClass="button-Red" OnClick="bbExit_Click" Text="結束" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-4Col" />
                <td class="ColWidth-4Col" />
                <td class="ColWidth-4Col" />
                <td class="ColWidth-4Col" />
            </tr>
        </table>
    </asp:Panel>
    
    </asp:Content>
