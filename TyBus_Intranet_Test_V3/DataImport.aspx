<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="DataImport.aspx.cs" Inherits="TyBus_Intranet_Test_V3.DataImport" %>

<asp:Content ID="DataImportForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">資料匯入</a>
    </div>
    <br />
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbTargetTable" runat="server" CssClass="text-Right-Blue" Text="轉入目的：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:DropDownList ID="ddlTargetTable" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnSelectedIndexChanged="ddlTargetTable_SelectedIndexChanged" Width="97%">
                        <asp:ListItem Text="" Value="" Selected="True" />
                        <asp:ListItem Text="轉入違規案件" Value="T001" />
                        <asp:ListItem Text="轉入肇事案件" Value="T002" />
                        <asp:ListItem Text="轉入維修項目" Value="T003" />
                        <asp:ListItem Text="轉入客服案件" Value="T004" />
                        <asp:ListItem Text="轉入保養預排" Value="T005" />
                        <asp:ListItem Text="轉入合約資料" Value="T006" />
                        <asp:ListItem Text="轉入路線對照" Value="T007" />
                        <asp:ListItem Text="更新不休假獎金資料" Value="T008" />
                        <asp:ListItem Text="匯入MOU津貼資料" Value="T009" />
                        <asp:ListItem Text="修正MOU津貼資料" Value="T010" />
                    </asp:DropDownList>
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col">
                    <asp:Label ID="lbSourceFile" runat="server" CssClass="text-Right-Blue" Text="來源檔案：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="3">
                    <asp:FileUpload ID="fuExcel" runat="server" CssClass="text-Left-Black" Width="95%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-9Col" colspan="2">
                    <asp:Label ID="lbCalYM" runat="server" CssClass="text-Right-Blue" Text="計算年月：" Width="30%" />
                    <asp:TextBox ID="eCalYear" runat="server" CssClass="text-Left-Black" Width="15%" />
                    <asp:Label ID="lbCalYear" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:TextBox ID="eCalMonth" runat="server" CssClass="text-Left-Black" Width="15%" />
                    <asp:Label ID="lbCalMonth" runat="server" CssClass="text-Left-Black" Text=" 月" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-9Col">
                    <asp:Button ID="bbImport" runat="server" CssClass="button-Black" Text="匯入" OnClick="bbImport_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-9Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-9Col" colspan="7">
                    <asp:Label ID="lbError" runat="server" CssClass="errorMessageText" />
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
    </asp:Panel>
    <asp:Panel ID="plShowError" runat="server" CssClass="ShowPanel">
        <asp:TextBox ID="eShowError" runat="server" CssClass="text-Left-Red" TextMode="MultiLine" Width="100%" Height="500px" />
    </asp:Panel>
    <asp:ObjectDataSource ID="odsCustomServiceImport" runat="server" InsertMethod="InsertQuery_Import" OldValuesParameterFormatString="original_{0}" SelectMethod="GetData_Import" TypeName="CustomServiceTableAdapters.dtCustomService_ImportTableAdapter">
        <InsertParameters>
            <asp:Parameter Name="ServiceNo" Type="String" />
            <asp:Parameter Name="BuildDate" Type="DateTime" />
            <asp:Parameter Name="BuildTime" Type="String" />
            <asp:Parameter Name="BuildMan" Type="String" />
            <asp:Parameter Name="ServiceType" Type="String" />
            <asp:Parameter Name="ServiceTypeB" Type="String" />
            <asp:Parameter Name="ServiceTypeC" Type="String" />
            <asp:Parameter Name="LinesNo" Type="String" />
            <asp:Parameter Name="Car_ID" Type="String" />
            <asp:Parameter Name="Driver" Type="String" />
            <asp:Parameter Name="DriverName" Type="String" />
            <asp:Parameter Name="BoardTime" Type="String" />
            <asp:Parameter Name="BoardStation" Type="String" />
            <asp:Parameter Name="GetoffTime" Type="String" />
            <asp:Parameter Name="GetoffStation" Type="String" />
            <asp:Parameter Name="ServiceNote" Type="String" />
            <asp:Parameter Name="IsNoContect" Type="Boolean" />
            <asp:Parameter Name="CivicName" Type="String" />
            <asp:Parameter Name="CivicTelNo" Type="String" />
            <asp:Parameter Name="CivicCellPhone" Type="String" />
            <asp:Parameter Name="CivicAddress" Type="String" />
            <asp:Parameter Name="CivicEMail" Type="String" />
            <asp:Parameter Name="AthorityDepNo" Type="String" />
            <asp:Parameter Name="AthorityDepNote" Type="String" />
            <asp:Parameter Name="IsReplied" Type="Boolean" />
            <asp:Parameter Name="Remark" Type="String" />
            <asp:Parameter Name="IsPending" Type="Boolean" />
            <asp:Parameter Name="AssignDate" Type="DateTime" />
            <asp:Parameter Name="AssignMan" Type="String" />
            <asp:Parameter Name="IsClosed" Type="Boolean" />
            <asp:Parameter Name="CloseDate" Type="DateTime" />
            <asp:Parameter Name="CloseMan" Type="String" />
        </InsertParameters>
    </asp:ObjectDataSource>
</asp:Content>
